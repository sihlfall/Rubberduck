using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting.Execution
{
    /// <summary>
    /// Shared execution policy for discovered and external plans. No declarations,
    /// parse trees, navigation, VBE objects, or Test Explorer models enter this class.
    /// The process-wide gate protects AssertHandler's static event and host globals.
    /// </summary>
    public sealed class TestExecutionCore
    {
        private static int _running;
        public static bool IsRunning => Volatile.Read(ref _running) != 0;

        public long Execute(TestExecutionPlan plan, ITestExecutionHost host, IFakesFactory fakesFactory,
            Func<bool> cancelled, Action pump, Action<ExecutionNotification> publish)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            plan.Validate();
            if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
                throw new InvalidOperationException("Another test run is already executing in this host.");
            var watch = Stopwatch.StartNew();
            try
            {
                publish(new ExecutionNotification { Kind = ExecutionNotificationKind.RunStarted });
                foreach (var module in plan.Modules)
                {
                    if (cancelled()) break;
                    var fakes = fakesFactory.Create();
                    using (var session = host.OpenModule(module))
                    {
                        try { RunHooks(session, module.ModuleInitialize); }
                        catch (COMException error)
                        {
                            Issue(publish, module, "moduleInitialize", error);
                            foreach (var test in module.Tests)
                                Complete(publish, module, test, new TestResult(TestOutcome.Unknown, AssertMessages.TestRunner_ModuleInitializeFailure));
                            // Preserve the full runner's established module-setup failure policy.
                            continue;
                        }
                        foreach (var test in module.Tests)
                        {
                            if (cancelled()) break;
                            publish(new ExecutionNotification { Kind = ExecutionNotificationKind.TestStarted, Module = module, Test = test });
                            if (test.Ignored)
                            {
                                Complete(publish, module, test, new TestResult(TestOutcome.Ignored));
                                continue;
                            }
                            try
                            {
                                fakes.StartTest();
                                try { RunHooks(session, module.TestInitialize); }
                                catch (COMException error)
                                {
                                    Issue(publish, module, "testInitialize", error, test);
                                    Complete(publish, module, test, new TestResult(TestOutcome.Inconclusive, AssertMessages.TestRunner_TestInitializeFailure));
                                    continue;
                                }
                                pump();
                                if (cancelled())
                                {
                                    Cleanup(session, module.TestCleanup, publish, module, "testCleanup", test);
                                    break;
                                }
                                Complete(publish, module, test, RunTest(session, test));
                                Cleanup(session, module.TestCleanup, publish, module, "testCleanup", test);
                            }
                            finally { fakes.StopTest(); }
                        }
                        try { RunHooks(session, module.ModuleCleanup); }
                        catch (COMException error)
                        {
                            Issue(publish, module, "moduleCleanup", error);
                            break;
                        }
                    }
                }
            }
            catch (Exception error)
            {
                publish(new ExecutionNotification { Kind = ExecutionNotificationKind.Issue, Phase = "run", Message = error.ToString() });
            }
            finally
            {
                watch.Stop();
                try
                {
                    publish(new ExecutionNotification { Kind = ExecutionNotificationKind.RunCompleted, Duration = watch.ElapsedMilliseconds, Cancelled = cancelled() });
                }
                finally { Interlocked.Exchange(ref _running, 0); }
            }
            return watch.ElapsedMilliseconds;
        }

        private static void RunHooks(ITestModuleSession session, IEnumerable<ExecutionProcedure> procedures)
        {
            foreach (var procedure in procedures) session.RunHook(procedure);
        }

        private static void Cleanup(ITestModuleSession session, IEnumerable<ExecutionProcedure> procedures,
            Action<ExecutionNotification> publish, ExecutionModule module, string phase, ExecutionTest test)
        {
            try { RunHooks(session, procedures); }
            catch (COMException error) { Issue(publish, module, phase, error, test); }
        }

        private static TestResult RunTest(ITestModuleSession session, ExecutionTest test)
        {
            long duration = 0;
            try
            {
                var assertions = new List<AssertCompletedEventArgs>();
                session.RunTest(test, (sender, assertion) => assertions.Add(assertion), out duration);
                var firstFailure = assertions.FirstOrDefault(assertion => assertion.Outcome != TestOutcome.Succeeded);
                return firstFailure == null
                    ? new TestResult(TestOutcome.Succeeded, duration: duration)
                    : new TestResult(firstFailure.Outcome, firstFailure.Message, duration);
            }
            catch (COMException) { return new TestResult(TestOutcome.Inconclusive, AssertMessages.TestRunner_ComException, duration); }
            catch (Exception) { return new TestResult(TestOutcome.Inconclusive, AssertMessages.TestRunner_ExceptionDuringRun, duration); }
        }

        private static void Complete(Action<ExecutionNotification> publish, ExecutionModule module, ExecutionTest test, TestResult result)
        {
            publish(new ExecutionNotification { Kind = ExecutionNotificationKind.TestCompleted, Module = module, Test = test, Result = result });
        }

        private static void Issue(Action<ExecutionNotification> publish, ExecutionModule module, string phase, Exception error, ExecutionTest test = null)
        {
            publish(new ExecutionNotification { Kind = ExecutionNotificationKind.Issue, Module = module, Test = test, Phase = phase, Message = error.ToString() });
        }
    }
}
