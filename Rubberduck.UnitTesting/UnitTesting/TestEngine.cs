using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using NLog;
using Rubberduck.UnitTesting.Execution;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    // FIXME litter logging around here
    internal class TestEngine : ITestEngine
    {
        protected static readonly ParserState[] AllowedRunStates = 
        {
            ParserState.Ready
        };

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly RubberduckParserState _state;
        private readonly IFakesFactory _fakesFactory;
        private readonly IVBEInteraction _declarationRunner;
        private readonly ITypeLibWrapperProvider _wrapperProvider;
        private readonly IUiDispatcher _uiDispatcher;
        private readonly IVBE _vbe;
        private readonly IProjectsProvider _projectsProvider;

        private Dictionary<TestMethod, TestOutcome> _knownOutcomes = new Dictionary<TestMethod, TestOutcome>();
        private List<TestMethod> _lastRun = new List<TestMethod>();
        private List<TestMethod> _tests;

        public IEnumerable<TestMethod> Tests => _tests;

        public IReadOnlyList<TestMethod> LastRunTests => _lastRun;

        public bool CanRun => AllowedRunStates.Contains(_state.Status) && _vbe.IsInDesignMode && !TestExecutionCore.IsRunning;
        public bool CanRepeatLastRun => _lastRun.Any();
        
        private bool _listening = true;

        public TestEngine(
            RubberduckParserState state, 
            IFakesFactory fakesFactory, 
            IVBEInteraction declarationRunner, 
            ITypeLibWrapperProvider wrapperProvider, 
            IUiDispatcher uiDispatcher,
            IVBE vbe,
            IProjectsProvider projectsProvider)
        {
            Debug.WriteLine("TestEngine created.");
            _state = state;
            _fakesFactory = fakesFactory;
            _declarationRunner = declarationRunner;
            _wrapperProvider = wrapperProvider;
            _uiDispatcher = uiDispatcher;
            _vbe = vbe;
            _projectsProvider = projectsProvider;

            _state.StateChanged += StateChangedHandler;
        }

        private void StateChangedHandler(object sender, ParserStateEventArgs e)
        {
            if (e.OldState == ParserState.Started)
            {
                _uiDispatcher.InvokeAsync(() => TestsRefreshStarted?.Invoke(this, EventArgs.Empty));
                return;
            }

            if (!CanRun || e.IsError)
            {
                _listening = true;
            }
            // CanRun returned true already, only refresh tests if we're not backed off
            else if (_listening && e.OldState != ParserState.Busy)
            {
                _listening = false;
                var updates = TestDiscovery.GetAllTests(_state).ToList();
                var run = new List<TestMethod>();
                var known = new Dictionary<TestMethod, TestOutcome>();

                foreach (var test in updates)
                {
                    var match = _lastRun.FirstOrDefault(ut => ut.Equals(test));
                    if (match != null)
                    {
                        run.Add(match);
                    }

                    if (_knownOutcomes.ContainsKey(test))
                    {
                        known.Add(test, _knownOutcomes[test]);
                    }
                }

                _tests = updates;
                _lastRun = run;
                _knownOutcomes = known;
                _uiDispatcher.InvokeAsync(() => TestsRefreshed?.Invoke(this, EventArgs.Empty));
            }
        }

        public event EventHandler<TestRunStartedEventArgs> TestRunStarted;
        public event EventHandler<TestStartedEventArgs> TestStarted;
        public event EventHandler<TestCompletedEventArgs> TestCompleted;
        public event EventHandler<TestRunCompletedEventArgs> TestRunCompleted;
        public event EventHandler TestsRefreshStarted;
        public event EventHandler TestsRefreshed;

        private void OnTestRunStarted(IReadOnlyList<TestMethod> tests)
        {
            CancellationRequested = false;
            TestRunStarted?.Invoke(this, new TestRunStartedEventArgs(tests));
            // This call is safe - OnTestRunStarted cannot be called from outside RD's context.
            _uiDispatcher.FlushMessageQueue(); 
        }

        private void OnTestStarted(TestMethod test)
        {           
            TestStarted?.Invoke(this, new TestStartedEventArgs(test));
            // This call is safe - OnTestStarted cannot be called from outside RD's context.
            _uiDispatcher.FlushMessageQueue();
        }

        private void OnTestCompleted(TestMethod test, TestResult result)
        {
            _lastRun.Add(test);
            _knownOutcomes.Add(test, result.Outcome);

            TestCompleted?.Invoke(this, new TestCompletedEventArgs(test, result));
            // This call is safe - OnTestCompleted cannot be called from outside RD's context.
            _uiDispatcher.FlushMessageQueue();
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            var queued = tests.ToList();
            if (!queued.Any()) return;
            if (!CanRun) throw new InvalidOperationException("The host is not ready for a test run.");

            foreach (var test in queued.Where(item => _knownOutcomes.ContainsKey(item)))
            {
                _knownOutcomes.Remove(test);
            }

            _uiDispatcher.InvokeAsync(() =>
            {
                if (!CanRun)
                {
                    TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0));
                    return;
                }
                OnTestRunStarted(queued);
                RunInternal(queued);
            });
        }

        public void RunByOutcome(TestOutcome outcome)
        {
            Run(_knownOutcomes.Where(test => test.Value == outcome).Select(test => test.Key));
        }

        public void RepeatLastRun()
        {
            Run(_lastRun);
        }

        private bool CancellationRequested { get; set; }

        public void RequestCancellation()
        {
            CancellationRequested = true;
        }

        protected virtual void RunInternal(IEnumerable<TestMethod> tests)
        {
            if (!CanRun)
            {
                TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0));
                return;
            }
            //We push the suspension to a background thread to avoid potential deadlocks if a parse is still running.
            Task.Run(() =>
            {
                var suspensionResult = _state.OnSuspendParser(this, AllowedRunStates, () => RunWhileSuspended(tests));

                //We have to log and swallow since we run as the top level code in a background thread.
                switch (suspensionResult.Outcome)
                {
                    case SuspensionOutcome.Completed:
                        return;
                    case SuspensionOutcome.Canceled:
                        Logger.Debug("Test execution canceled.");
                        return;
                    default:
                        Logger.Warn($"Test execution failed with suspension outcome {suspensionResult.Outcome}.");
                        if (suspensionResult.EncounteredException != null)
                        {
                            Logger.Error(suspensionResult.EncounteredException);
                        }

                        _uiDispatcher.InvokeAsync(() => TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0)));
                        return;
                }
            });
        }

        private void EnsureRubberduckIsReferencedForEarlyBoundTests()
        {
            var projectIdsOfMembersUsingAddInLibrary = _state.DeclarationFinder.AllUserDeclarations
                .Where(member => member.AsTypeName == "Rubberduck.PermissiveAssertClass"
                                 || member.AsTypeName == "Rubberduck.AssertClass")
                .Select(member => member.ProjectId)
                .Distinct();

            var projectsUsingAddInLibrary = projectIdsOfMembersUsingAddInLibrary
                .Select(projectId => _projectsProvider.Project(projectId))
                .Where(project => project != null);

            foreach (var project in projectsUsingAddInLibrary)
            {
                _declarationRunner.EnsureProjectReferencesUnitTesting(project);
            }
        }

        protected void RunWhileSuspended(IEnumerable<TestMethod> tests)
        {
            //Running the tests has to be done on the UI thread, so we push the task to it from within suspension of the parser.
            //We have to wait for the completion to make sure that the suspension only ends after tests have been completed.
            var testTask = _uiDispatcher.StartTask(() => RunWhileSuspendedOnUiThread(tests));
            testTask.Wait();
        }

        private void RunWhileSuspendedOnUiThread(IEnumerable<TestMethod> tests)
        {
            var testMethods = tests.ToList();
            if (!testMethods.Any()) return;
            _lastRun.Clear();
            try
            {
                EnsureRubberduckIsReferencedForEarlyBoundTests();
            }
            catch (InvalidOperationException error)
            {
                Logger.Warn(error);
                foreach (var test in testMethods)
                    OnTestCompleted(test, new TestResult(TestOutcome.Failed, AssertMessages.Prerequisite_EarlyBindingReferenceMissing));
                TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0));
                return;
            }

            var plan = RubberduckTestPlanAdapter.Create(testMethods, _state);
            var host = new DiscoveredTestExecutionHost(_declarationRunner, _wrapperProvider);
            try
            {
                new TestExecutionCore().Execute(plan, host, _fakesFactory, () => CancellationRequested,
                    _uiDispatcher.FlushMessageQueue, notification =>
                    {
                        switch (notification.Kind)
                        {
                            case ExecutionNotificationKind.TestStarted:
                                OnTestStarted((TestMethod)notification.Test.Tag);
                                break;
                            case ExecutionNotificationKind.TestCompleted:
                                OnTestCompleted((TestMethod)notification.Test.Tag, notification.Result);
                                break;
                            case ExecutionNotificationKind.Issue:
                                Logger.Warn(notification.Message);
                                break;
                            case ExecutionNotificationKind.RunCompleted:
                                CancellationRequested = false;
                                TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(notification.Duration));
                                break;
                        }
                    });
            }
            catch (Exception error)
            {
                Logger.Error(error, "Test plan execution failed.");
                CancellationRequested = false;
                TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0));
            }
        }
    }
}
