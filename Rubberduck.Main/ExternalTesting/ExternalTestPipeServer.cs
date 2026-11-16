using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Rubberduck.UnitTesting.Execution;

namespace Rubberduck.ExternalTesting
{
    /// <summary>Opt-in, current-user-only IPC. No COM/VBE objects are touched on the transport threads.</summary>
    public sealed class ExternalTestPipeServer : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly IExternalTestExecutionService _service;
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private readonly ConcurrentDictionary<NamedPipeServerStream, byte> _connections = new ConcurrentDictionary<NamedPipeServerStream, byte>();
        private readonly SemaphoreSlim _slots = new SemaphoreSlim(8);
        private int _started;
        public string PipeName { get; }

        public ExternalTestPipeServer(IExternalTestExecutionService service, string pipeName = null)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            PipeName = pipeName ?? "Rubberduck.Tests." + Process.GetCurrentProcess().Id;
        }

        public void Start()
        {
            if (Interlocked.CompareExchange(ref _started, 1, 0) != 0) throw new InvalidOperationException("Pipe server already started.");
            _ = Task.Run(AcceptAsync);
            Logger.Info("External test control enabled on pipe " + PipeName);
        }

        private NamedPipeServerStream CreatePipe()
        {
            var security = new PipeSecurity();
            security.SetAccessRuleProtection(true, false);
            using (var identity = WindowsIdentity.GetCurrent())
                security.AddAccessRule(new PipeAccessRule(identity.User, PipeAccessRights.FullControl, AccessControlType.Allow));
            return new NamedPipeServerStream(PipeName, PipeDirection.InOut, 8, PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous, 65536, 65536, security);
        }

        private async Task AcceptAsync()
        {
            try
            {
                while (!_shutdown.IsCancellationRequested)
                {
                    await _slots.WaitAsync(_shutdown.Token).ConfigureAwait(false);
                    NamedPipeServerStream pipe;
                    try { pipe = CreatePipe(); }
                    catch { _slots.Release(); throw; }
                    _connections.TryAdd(pipe, 0);
                    try { await pipe.WaitForConnectionAsync(_shutdown.Token).ConfigureAwait(false); }
                    catch { Close(pipe); throw; }
                    // A separate connection can cancel a run while its event stream is open.
                    _ = Task.Run(() => HandleAsync(pipe));
                }
            }
            catch (Exception error)
            {
                if (!_shutdown.IsCancellationRequested) Logger.Error(error, "External test pipe listener stopped.");
            }
        }

        private async Task HandleAsync(NamedPipeServerStream pipe)
        {
            string runId = null;
            try
            {
                using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(_shutdown.Token))
                using (deadline.Token.Register(() => Close(pipe)))
                {
                    deadline.CancelAfter(TimeSpan.FromSeconds(10));
                    var request = ExternalTestProtocol.Deserialize<ExternalTestRequest>(await ReadFrameAsync(pipe, deadline.Token).ConfigureAwait(false));
                    runId = request.RunId;
                    if (request.Version != ExternalTestProtocol.Version) throw new InvalidDataException("Unsupported protocol version.");
                    switch (request.Command)
                    {
                        case "describe":
                            await WriteFrameAsync(pipe, await _service.DescribeAsync(deadline.Token).ConfigureAwait(false), deadline.Token).ConfigureAwait(false);
                            break;
                        case "cancel":
                            if (string.IsNullOrWhiteSpace(runId)) throw new InvalidDataException("runId is required.");
                            await WriteFrameAsync(pipe, new ExternalTestEvent
                            {
                                Event = "cancel_requested", RunId = runId,
                                Cancelled = _service.Cancel(runId)
                            }, deadline.Token).ConfigureAwait(false);
                            break;
                        case "run":
                            // Run has its own cooperative timeout; the request-read deadline is complete.
                            deadline.CancelAfter(Timeout.Infinite);
                            await StreamRunAsync(pipe, request, deadline.Token).ConfigureAwait(false);
                            break;
                        default: throw new InvalidDataException("Unknown command.");
                    }
                }
            }
            catch (Exception error)
            {
                if (!_shutdown.IsCancellationRequested)
                {
                    Logger.Warn(error, "External test request failed.");
                    try
                    {
                        using (var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(2)))
                            await WriteFrameAsync(pipe, new ExternalTestEvent { Event = "error", RunId = runId, Message = error.Message }, deadline.Token).ConfigureAwait(false);
                    }
                    catch (Exception) { /* A disconnected client cannot receive an error. */ }
                }
            }
            finally { Close(pipe); }
        }

        private async Task StreamRunAsync(NamedPipeServerStream pipe, ExternalTestRequest request, CancellationToken shutdown)
        {
            using (var cancellation = CancellationTokenSource.CreateLinkedTokenSource(shutdown))
            using (cancellation.Token.Register(() => Close(pipe)))
            using (var events = new BlockingCollection<ExternalTestEvent>(4096))
            {
                ExternalTestEvent terminal = null;
                var writer = Task.Run(async () =>
                {
                    foreach (var item in events.GetConsumingEnumerable(cancellation.Token))
                        await WriteFrameAsync(pipe, item, cancellation.Token).ConfigureAwait(false);
                });
                var execution = _service.RunAsync(request, item =>
                {
                    // Publish terminal success only after the service has released its
                    // run reservation and parser suspension.
                    if (item.Event == "run_completed") { terminal = item; return; }
                    // Never block the host UI thread behind a slow or disconnected client.
                    if (!events.TryAdd(item))
                    {
                        cancellation.Cancel();
                        throw new IOException("External test event queue is full.");
                    }
                }, cancellation.Token);
                // One command per connection; EOF or extra input ends/cancels this run.
                var disconnect = pipe.ReadAsync(new byte[1], 0, 1, cancellation.Token);
                try
                {
                    var first = await Task.WhenAny(execution, writer, disconnect).ConfigureAwait(false);
                    if (first != execution) cancellation.Cancel();
                    await execution.ConfigureAwait(false);
                    if (terminal == null) throw new InvalidDataException("Execution returned without a terminal result.");
                    if (!events.TryAdd(terminal)) throw new IOException("External test event queue is full.");
                }
                catch (Exception error) when (!cancellation.IsCancellationRequested)
                {
                    // Keep preflight errors on the run stream before cancellation
                    // closes the pipe in the finally block.
                    if (!events.TryAdd(new ExternalTestEvent { Event = "error", RunId = request.RunId, Message = error.Message }))
                        throw;
                }
                finally
                {
                    events.CompleteAdding();
                    try { await writer.ConfigureAwait(false); }
                    finally
                    {
                        cancellation.Cancel();
                        // Framework pipe reads may not honor cancellation until the pipe
                        // closes. HandleAsync closes it after sending any terminal error.
                        _ = disconnect.ContinueWith(task => { var ignored = task.Exception; },
                            CancellationToken.None, TaskContinuationOptions.OnlyOnFaulted, TaskScheduler.Default);
                    }
                }
            }
        }

        public static async Task<byte[]> ReadFrameAsync(Stream stream, CancellationToken token)
        {
            var header = new byte[4];
            await ReadExactlyAsync(stream, header, token).ConfigureAwait(false);
            var length = header[0] | header[1] << 8 | header[2] << 16 | header[3] << 24;
            if (length < 1 || length > ExternalTestProtocol.MaximumFrameBytes) throw new InvalidDataException("Invalid frame size.");
            var payload = new byte[length];
            await ReadExactlyAsync(stream, payload, token).ConfigureAwait(false);
            return payload;
        }

        private static async Task ReadExactlyAsync(Stream stream, byte[] buffer, CancellationToken token)
        {
            var offset = 0;
            while (offset < buffer.Length)
            {
                var read = await stream.ReadAsync(buffer, offset, buffer.Length - offset, token).ConfigureAwait(false);
                if (read == 0) throw new EndOfStreamException("Incomplete protocol frame.");
                offset += read;
            }
        }

        public static async Task WriteFrameAsync(Stream stream, ExternalTestEvent value, CancellationToken token)
        {
            var payload = ExternalTestProtocol.Serialize(value);
            if (payload.Length > ExternalTestProtocol.MaximumFrameBytes) throw new InvalidDataException("Response exceeds maximum frame size.");
            var length = payload.Length;
            var header = new[] { (byte)length, (byte)(length >> 8), (byte)(length >> 16), (byte)(length >> 24) };
            await stream.WriteAsync(header, 0, header.Length, token).ConfigureAwait(false);
            await stream.WriteAsync(payload, 0, payload.Length, token).ConfigureAwait(false);
            await stream.FlushAsync(token).ConfigureAwait(false);
        }

        private void Close(NamedPipeServerStream pipe)
        {
            if (_connections.TryRemove(pipe, out _))
            {
                pipe.Dispose();
                _slots.Release();
            }
        }

        public void Dispose()
        {
            _shutdown.Cancel();
            foreach (var pipe in _connections.Keys) Close(pipe);
            // Do not wait on an executing host UI task from host shutdown.
        }
    }
}
