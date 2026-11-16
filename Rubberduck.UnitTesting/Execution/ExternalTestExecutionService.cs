using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting.Execution
{
    /// <summary>Full-package adapter: parser suspension coordinates the host; discovery comes entirely from the request.</summary>
    public sealed class ExternalTestExecutionService : IExternalTestExecutionService
    {
        private static readonly ParserState[] StableStates =
        {
            ParserState.Pending, ParserState.Ready, ParserState.Error, ParserState.ResolverError, ParserState.UnexpectedError
        };
        private readonly RubberduckParserState _state;
        private readonly IUiDispatcher _dispatcher;
        private readonly IVBE _vbe;
        private readonly IProjectsRepository _projects;
        private readonly ITypeLibWrapperProvider _wrappers;
        private readonly IVBETypeLibsAPI _typeLibs;
        private readonly IHostProcedureInvoker _invoker;
        private readonly IFakesFactory _fakes;
        private readonly object _sync = new object();
        private string _runId;
        private CancellationTokenSource _cancellation;

        public ExternalTestExecutionService(RubberduckParserState state, IUiDispatcher dispatcher, IVBE vbe,
            IProjectsRepository projects, ITypeLibWrapperProvider wrappers, IVBETypeLibsAPI typeLibs,
            IHostProcedureInvoker invoker, IFakesFactory fakes)
        {
            _state = state; _dispatcher = dispatcher; _vbe = vbe; _projects = projects;
            _wrappers = wrappers; _typeLibs = typeLibs; _invoker = invoker; _fakes = fakes;
        }

        public Task<ExternalTestEvent> DescribeAsync(CancellationToken token)
        {
            return _dispatcher.StartTask(() =>
            {
                token.ThrowIfCancellationRequested();
                if (!_vbe.IsInDesignMode || TestExecutionCore.IsRunning)
                    throw new InvalidOperationException("Host must be idle and in design mode.");
                _projects.Refresh();
                return new ExternalTestEvent
                {
                    Event = "description", HostKind = _vbe.Kind.ToString(), Bitness = IntPtr.Size * 8,
                    Projects = _projects.Projects().Where(project => project.Project.Protection == ProjectProtection.Unprotected)
                        .Select(project => Snapshot(project.Project)).ToList()
                };
            }, token);
        }

        public bool Cancel(string runId)
        {
            lock (_sync)
            {
                if (_cancellation == null || !string.Equals(runId, _runId, StringComparison.Ordinal)) return false;
                _cancellation.Cancel();
                return true;
            }
        }

        public async Task RunAsync(ExternalTestRequest request, Action<ExternalTestEvent> publish, CancellationToken token)
        {
            if (request == null || request.Version != ExternalTestProtocol.Version || request.Command != "run")
                throw new ArgumentException("Unsupported run request.");
            if (string.IsNullOrWhiteSpace(request.RunId) || request.RunId.Length > 128)
                throw new ArgumentException("A runId of at most 128 characters is required.");
            if (request.Plan == null) throw new ArgumentException("A plan is required.");
            request.Plan.Validate();
            if (request.TimeoutSeconds < 1 || request.TimeoutSeconds > 3600)
                throw new ArgumentException("timeoutSeconds must be between 1 and 3600.");
            var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token);
            lock (_sync)
            {
                if (_cancellation != null) { cancellation.Dispose(); throw new InvalidOperationException("An external run is already pending."); }
                _runId = request.RunId;
                _cancellation = cancellation;
            }
            cancellation.CancelAfter(TimeSpan.FromSeconds(request.TimeoutSeconds));
            try
            {
                await Task.Run(() =>
                {
                    var suspension = _state.OnSuspendParser(this, StableStates, () =>
                        _dispatcher.StartTask(() =>
                        {
                            cancellation.Token.ThrowIfCancellationRequested();
                            ValidateHost(request);
                            new TestExecutionCore().Execute(request.Plan, new ExternalHost(_wrappers, _invoker), _fakes,
                                () => cancellation.IsCancellationRequested, _dispatcher.FlushMessageQueue,
                                notification => publish(ToEvent(request.RunId, notification)));
                        }, cancellation.Token).GetAwaiter().GetResult(), 10000);
                    if (suspension.Outcome != SuspensionOutcome.Completed)
                        throw new InvalidOperationException("Host execution could not acquire parser suspension: " + suspension.Outcome, suspension.EncounteredException);
                }).ConfigureAwait(false);
            }
            finally
            {
                lock (_sync) { _runId = null; _cancellation = null; }
                cancellation.Dispose();
            }
        }

        private void ValidateHost(ExternalTestRequest request)
        {
            if (!_vbe.IsInDesignMode || TestExecutionCore.IsRunning)
                throw new InvalidOperationException("Host must be idle and in design mode.");
            if (request.Bitness != IntPtr.Size * 8 || request.HostKind != _vbe.Kind.ToString())
                throw new ArgumentException("The plan's host kind or bitness does not match this host.");
            if (request.Projects == null) throw new ArgumentException("Project source snapshots are required.");
            _projects.Refresh();
            var expected = request.Projects.ToDictionary(project => project.ProjectId, StringComparer.Ordinal);
            var selectedProjects = request.Plan.Modules.Select(module => module.ProjectId).Distinct().ToList();
            if (expected.Count != selectedProjects.Count) throw new ArgumentException("Supply exactly the selected projects' snapshots.");
            foreach (var id in selectedProjects)
            {
                var project = _projects.Project(id);
                if (project == null || project.Protection != ProjectProtection.Unprotected)
                    throw new InvalidOperationException("Project unavailable or protected: " + id);
                if (!expected.TryGetValue(id, out var snapshot)) throw new ArgumentException("Missing project snapshot.");
                ValidateSnapshot(snapshot, Snapshot(project));
                using (var components = project.VBComponents)
                {
                    foreach (var module in request.Plan.Modules.Where(module => module.ProjectId == id))
                    using (var component = components[module.Name])
                    {
                        if (component == null || component.IsWrappingNullReference || component.Type != ComponentType.StandardModule)
                            throw new ArgumentException("Test modules must be existing standard modules: " + module.Name);
                    }
                }
            }
            // All snapshots are checked before changing references or compiling any project.
            foreach (var id in selectedProjects)
            {
                var project = _projects.Project(id);
                _invoker.EnsureProjectReferencesUnitTesting(project);
                if (!_typeLibs.CompileProject(project)) throw new InvalidOperationException("Host compilation failed: " + id);
            }
        }

        private ExternalProjectSnapshot Snapshot(IVBProject project)
        {
            var sources = new List<ExternalSourceSnapshot>();
            using (var components = project.VBComponents)
            foreach (var component in components)
            using (component)
            {
                if (!component.HasCodeModule) continue;
                using (var code = component.CodeModule)
                    sources.Add(new ExternalSourceSnapshot { Module = component.Name, Sha256 = ExternalTestProtocol.SourceHash(code.Content()) });
            }
            return new ExternalProjectSnapshot
            {
                ProjectId = project.ProjectId, Name = project.Name, Path = project.FileName,
                ConditionalCompilation = _typeLibs.GetProjectConditionalCompilationArgsRaw(project) ?? string.Empty,
                Sources = sources
            };
        }

        public static void ValidateSnapshot(ExternalProjectSnapshot expected, ExternalProjectSnapshot actual)
        {
            if (expected == null || expected.ConditionalCompilation != actual.ConditionalCompilation || expected.Sources == null)
                throw new ArgumentException("Conditional compilation or source snapshot does not match the host.");
            var sources = expected.Sources.ToDictionary(source => source.Module, StringComparer.OrdinalIgnoreCase);
            if (sources.Count != actual.Sources.Count) throw new ArgumentException("Source component set has changed.");
            foreach (var source in actual.Sources)
                if (!sources.TryGetValue(source.Module, out var match) || !string.Equals(match.Sha256, source.Sha256, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("Source has changed: " + source.Module);
        }

        private static ExternalTestEvent ToEvent(string runId, ExecutionNotification notification)
        {
            return new ExternalTestEvent
            {
                RunId = runId,
                Event = notification.Kind == ExecutionNotificationKind.RunStarted ? "run_started"
                    : notification.Kind == ExecutionNotificationKind.TestStarted ? "test_started"
                    : notification.Kind == ExecutionNotificationKind.TestCompleted ? "test_completed"
                    : notification.Kind == ExecutionNotificationKind.RunCompleted ? "run_completed" : "issue",
                Module = notification.Module?.Name, TestId = notification.Test?.Id,
                Status = notification.Kind == ExecutionNotificationKind.TestCompleted ? notification.Result.Outcome.ToString() : null,
                Message = notification.Kind == ExecutionNotificationKind.TestCompleted ? notification.Result.Output : notification.Message,
                Phase = notification.Phase,
                DurationMs = notification.Kind == ExecutionNotificationKind.TestCompleted ? notification.Result.Duration : notification.Duration,
                Cancelled = notification.Cancelled
            };
        }

        private sealed class ExternalHost : ITestExecutionHost
        {
            private readonly ITypeLibWrapperProvider _wrappers;
            private readonly IHostProcedureInvoker _invoker;
            public ExternalHost(ITypeLibWrapperProvider wrappers, IHostProcedureInvoker invoker) { _wrappers = wrappers; _invoker = invoker; }
            public ITestModuleSession OpenModule(ExecutionModule module) => new Session(module.Name, _wrappers.TypeLibWrapperFromProject(module.ProjectId), _invoker);

            private sealed class Session : ITestModuleSession
            {
                private readonly string _module;
                private readonly ITypeLibWrapper _typeLib;
                private readonly IHostProcedureInvoker _invoker;
                public Session(string module, ITypeLibWrapper typeLib, IHostProcedureInvoker invoker) { _module = module; _typeLib = typeLib; _invoker = invoker; }
                public void RunHook(ExecutionProcedure procedure) => _invoker.RunProcedure(_typeLib, _module, procedure.Name);
                public void RunTest(ExecutionTest test, EventHandler<AssertCompletedEventArgs> assertions, out long duration)
                    => _invoker.RunTestProcedure(_typeLib, _module, test.Procedure, assertions, out duration);
                public void Dispose() => _typeLib.Dispose();
            }
        }
    }
}
