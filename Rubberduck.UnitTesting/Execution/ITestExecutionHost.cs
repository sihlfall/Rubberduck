using System;

namespace Rubberduck.UnitTesting.Execution
{
    public interface ITestExecutionHost
    {
        ITestModuleSession OpenModule(ExecutionModule module);
    }

    public interface ITestModuleSession : IDisposable
    {
        void RunHook(ExecutionProcedure procedure);
        void RunTest(ExecutionTest test, EventHandler<AssertCompletedEventArgs> assertions, out long duration);
    }

    public enum ExecutionNotificationKind { RunStarted, TestStarted, TestCompleted, Issue, RunCompleted }

    public sealed class ExecutionNotification
    {
        public ExecutionNotificationKind Kind { get; set; }
        public ExecutionModule Module { get; set; }
        public ExecutionTest Test { get; set; }
        public TestResult Result { get; set; }
        public string Phase { get; set; }
        public string Message { get; set; }
        public long Duration { get; set; }
        public bool Cancelled { get; set; }
    }
}
