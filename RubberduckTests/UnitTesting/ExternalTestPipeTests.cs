using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.ExternalTesting;
using Rubberduck.UnitTesting.Execution;

namespace RubberduckTests.UnitTesting
{
    [TestFixture, NonParallelizable]
    public class ExternalTestPipeTests
    {
        private sealed class Service : IExternalTestExecutionService
        {
            private CancellationTokenSource _run;
            private string _id;
            public bool RejectRuns;
            public readonly TaskCompletionSource<bool> Finished = new TaskCompletionSource<bool>();
            public Task<ExternalTestEvent> DescribeAsync(CancellationToken token)
                => Task.FromResult(new ExternalTestEvent { Event = "description", HostKind = "test", Bitness = 64 });
            public async Task RunAsync(ExternalTestRequest request, Action<ExternalTestEvent> publish, CancellationToken token)
            {
                if (RejectRuns) throw new ArgumentException("fixture preflight mismatch");
                using (var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token))
                {
                    _run = cancellation; _id = request.RunId;
                    publish(new ExternalTestEvent { Event = "run_started", RunId = _id });
                    try { await Task.Delay(Timeout.Infinite, cancellation.Token); }
                    catch (OperationCanceledException) { }
                    publish(new ExternalTestEvent { Event = "run_completed", RunId = _id, Cancelled = true });
                    _run = null;
                    Finished.TrySetResult(true);
                }
            }
            public bool Cancel(string id)
            {
                if (id != _id || _run == null) return false;
                _run.Cancel(); return true;
            }
        }

        private static async Task Send(NamedPipeClientStream pipe, ExternalTestRequest request)
        {
            var bytes = ExternalTestProtocol.Serialize(request);
            var header = BitConverter.GetBytes(bytes.Length);
            await pipe.WriteAsync(header, 0, header.Length);
            await pipe.WriteAsync(bytes, 0, bytes.Length);
            await pipe.FlushAsync();
        }

        private static async Task<ExternalTestEvent> Read(Stream pipe)
            => ExternalTestProtocol.Deserialize<ExternalTestEvent>(await ExternalTestPipeServer.ReadFrameAsync(pipe, CancellationToken.None));

        [Test, Timeout(10000)]
        public async Task RealPipeDescribesHostAndRejectsUnsupportedProtocol()
        {
            var name = "Rubberduck.Tests.Test." + Guid.NewGuid().ToString("N");
            using (var server = new ExternalTestPipeServer(new Service(), name))
            {
                server.Start();
                using (var client = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
                {
                    await client.ConnectAsync(5000);
                    await Send(client, new ExternalTestRequest { Version = 1, Command = "describe" });
                    Assert.That((await Read(client)).HostKind, Is.EqualTo("test"));
                }
                using (var client = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
                {
                    await client.ConnectAsync(5000);
                    await Send(client, new ExternalTestRequest { Version = 99, Command = "describe" });
                    Assert.That((await Read(client)).Event, Is.EqualTo("error"));
                }
            }
        }

        [Test, Timeout(10000)]
        public async Task SeparateConnectionCancelsOnlyTheRequestedRunAndStreamsCompletion()
        {
            var name = "Rubberduck.Tests.Test." + Guid.NewGuid().ToString("N");
            var service = new Service();
            using (var server = new ExternalTestPipeServer(service, name))
            using (var run = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
            {
                server.Start();
                await run.ConnectAsync(5000);
                await Send(run, new ExternalTestRequest { Version = 1, Command = "run", RunId = "chosen" });
                Assert.That((await Read(run)).Event, Is.EqualTo("run_started"));
                foreach (var id in new[] { "wrong", "chosen" })
                using (var cancel = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
                {
                    await cancel.ConnectAsync(5000);
                    await Send(cancel, new ExternalTestRequest { Version = 1, Command = "cancel", RunId = id });
                    Assert.That((await Read(cancel)).Cancelled, Is.EqualTo(id == "chosen"));
                }
                var completed = await Read(run);
                Assert.That(completed.Event, Is.EqualTo("run_completed"));
                Assert.That(completed.Cancelled, Is.True);
            }
        }

        [Test, Timeout(10000)]
        public async Task PreflightFailureReturnsAnErrorFrameInsteadOfBareEof()
        {
            var name = "Rubberduck.Tests.Test." + Guid.NewGuid().ToString("N");
            using (var server = new ExternalTestPipeServer(new Service { RejectRuns = true }, name))
            using (var client = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
            {
                server.Start();
                await client.ConnectAsync(5000);
                await Send(client, new ExternalTestRequest { Version = 1, Command = "run", RunId = "rejected" });
                var error = await Read(client);
                Assert.That(error.Event, Is.EqualTo("error"));
                Assert.That(error.RunId, Is.EqualTo("rejected"));
                Assert.That(error.Message, Does.Contain("preflight mismatch"));
            }
        }

        [Test, Timeout(10000)]
        public async Task DisconnectRequestsCancellation()
        {
            var name = "Rubberduck.Tests.Test." + Guid.NewGuid().ToString("N");
            var service = new Service();
            using (var server = new ExternalTestPipeServer(service, name))
            {
                server.Start();
                using (var run = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
                {
                    await run.ConnectAsync(5000);
                    await Send(run, new ExternalTestRequest { Version = 1, Command = "run", RunId = "disconnect" });
                    await Read(run);
                }
                Assert.That(await Task.WhenAny(service.Finished.Task, Task.Delay(5000)), Is.SameAs(service.Finished.Task));
            }
        }

        [TestCase(0), TestCase(-1), TestCase(1048577)]
        public void FrameReaderRejectsInvalidLengths(int length)
        {
            using (var input = new MemoryStream(BitConverter.GetBytes(length)))
                Assert.ThrowsAsync<InvalidDataException>(async () => await ExternalTestPipeServer.ReadFrameAsync(input, CancellationToken.None));
        }

        [Test]
        public void FrameReaderRejectsTruncatedPayload()
        {
            using (var input = new MemoryStream(new byte[] { 4, 0, 0, 0, 123 }))
                Assert.ThrowsAsync<EndOfStreamException>(async () => await ExternalTestPipeServer.ReadFrameAsync(input, CancellationToken.None));
        }
    }
}
