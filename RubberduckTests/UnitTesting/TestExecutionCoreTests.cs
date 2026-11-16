using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using NUnit.Framework;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.Execution;

namespace RubberduckTests.UnitTesting
{
    [TestFixture, NonParallelizable]
    public class TestExecutionCoreTests
    {
        private static TestExecutionPlan Plan()
        {
            return new TestExecutionPlan { Modules = new List<ExecutionModule>
            {
                new ExecutionModule
                {
                    ProjectId = "project", Name = "Tests",
                    ModuleInitialize = new List<ExecutionProcedure> { new ExecutionProcedure { Name = "ModuleBefore" } },
                    ModuleCleanup = new List<ExecutionProcedure> { new ExecutionProcedure { Name = "ModuleAfter" } },
                    TestInitialize = new List<ExecutionProcedure> { new ExecutionProcedure { Name = "Before" } },
                    TestCleanup = new List<ExecutionProcedure> { new ExecutionProcedure { Name = "After" } },
                    Tests = new List<ExecutionTest>
                    {
                        new ExecutionTest { Id = "one", Procedure = "One" },
                        new ExecutionTest { Id = "two", Procedure = "Two" }
                    }
                }
            } };
        }

        private sealed class Host : ITestExecutionHost, ITestModuleSession, IFakesFactory, IFakes
        {
            public readonly List<string> Trace = new List<string>();
            public string FailingHook;
            public Exception TestError;
            public readonly List<AssertCompletedEventArgs> Assertions = new List<AssertCompletedEventArgs>();
            public ITestModuleSession OpenModule(ExecutionModule module) { Trace.Add("open:" + module.Name); return this; }
            public void RunHook(ExecutionProcedure procedure)
            {
                Trace.Add(procedure.Name);
                if (procedure.Name == FailingHook) throw new COMException("fixture error");
            }
            public void RunTest(ExecutionTest test, EventHandler<AssertCompletedEventArgs> assertions, out long duration)
            {
                Trace.Add(test.Procedure);
                duration = 7;
                if (TestError != null) throw TestError;
                foreach (var assertion in Assertions) assertions(this, assertion);
            }
            public void Dispose() { Trace.Add("dispose"); }
            public IFakes Create() { Trace.Add("fakes"); return this; }
            public void StartTest() { Trace.Add("startFake"); }
            public void StopTest() { Trace.Add("stopFake"); }
        }

        private static List<ExecutionNotification> Execute(TestExecutionPlan plan, Host host, Func<bool> cancelled = null, Action pump = null)
        {
            var notifications = new List<ExecutionNotification>();
            new TestExecutionCore().Execute(plan, host, host, cancelled ?? (() => false), pump ?? (() => { }), notification =>
            {
                notifications.Add(notification);
                if (notification.Kind == ExecutionNotificationKind.TestCompleted) host.Trace.Add("result:" + notification.Test.Id);
            });
            return notifications;
        }

        [Test]
        public void SelectedTestsShareOneModuleLifecycleAndPublishBeforeCleanup()
        {
            var host = new Host();
            var events = Execute(Plan(), host);
            CollectionAssert.AreEqual(new[] {
                "fakes", "open:Tests", "ModuleBefore", "startFake", "Before", "One", "result:one", "After", "stopFake",
                "startFake", "Before", "Two", "result:two", "After", "stopFake", "ModuleAfter", "dispose"
            }, host.Trace);
            Assert.That(events.First().Kind, Is.EqualTo(ExecutionNotificationKind.RunStarted));
            Assert.That(events.Last().Kind, Is.EqualTo(ExecutionNotificationKind.RunCompleted));
            Assert.That(events.Where(item => item.Kind == ExecutionNotificationKind.TestCompleted).All(item => item.Result.Outcome == TestOutcome.Succeeded));
        }

        [Test]
        public void IgnoredTestsSkipPerTestHooksButRetainModuleHooks()
        {
            var plan = Plan();
            plan.Modules[0].Tests.ForEach(test => test.Ignored = true);
            var host = new Host();
            var events = Execute(plan, host);
            CollectionAssert.AreEqual(new[] { "fakes", "open:Tests", "ModuleBefore", "result:one", "result:two", "ModuleAfter", "dispose" }, host.Trace);
            Assert.That(events.Count(item => item.Kind == ExecutionNotificationKind.TestCompleted && item.Result.Outcome == TestOutcome.Ignored), Is.EqualTo(2));
        }

        [Test]
        public void BodyUsesFirstNonsuccessfulAssertionAndMilliseconds()
        {
            var host = new Host();
            host.Assertions.Add(new AssertCompletedEventArgs(TestOutcome.Succeeded));
            host.Assertions.Add(new AssertCompletedEventArgs(TestOutcome.Inconclusive, "first"));
            host.Assertions.Add(new AssertCompletedEventArgs(TestOutcome.Failed, "second"));
            var result = Execute(Plan(), host).First(item => item.Kind == ExecutionNotificationKind.TestCompleted).Result;
            Assert.That(result.Outcome, Is.EqualTo(TestOutcome.Inconclusive));
            Assert.That(result.Output, Is.EqualTo("first"));
            Assert.That(result.Duration, Is.EqualTo(7));
        }

        [Test]
        public void ModuleSetupFailureBlocksBodiesAndReportsIssue()
        {
            var host = new Host { FailingHook = "ModuleBefore" };
            var events = Execute(Plan(), host);
            Assert.That(events.Count(item => item.Kind == ExecutionNotificationKind.TestCompleted && item.Result.Outcome == TestOutcome.Unknown), Is.EqualTo(2));
            Assert.That(events.Single(item => item.Kind == ExecutionNotificationKind.Issue).Phase, Is.EqualTo("moduleInitialize"));
            Assert.That(host.Trace, Does.Not.Contain("One").And.Not.Contain("ModuleAfter"));
            Assert.That(host.Trace.Last(), Is.EqualTo("dispose"));
        }

        [Test]
        public void TestSetupFailureSkipsBodyAndCleanupAndStopsFakes()
        {
            var host = new Host { FailingHook = "Before" };
            var events = Execute(Plan(), host);
            Assert.That(events.Count(item => item.Kind == ExecutionNotificationKind.TestCompleted && item.Result.Outcome == TestOutcome.Inconclusive), Is.EqualTo(2));
            Assert.That(host.Trace, Does.Not.Contain("One").And.Not.Contain("After"));
            Assert.That(host.Trace.Count(item => item == "stopFake"), Is.EqualTo(2));
            Assert.That(host.Trace, Contains.Item("ModuleAfter"));
        }

        [Test]
        public void TestCleanupFailureRemainsSeparateAndAllowsNextBody()
        {
            var host = new Host { FailingHook = "After" };
            var events = Execute(Plan(), host);
            Assert.That(host.Trace, Contains.Item("Two"));
            Assert.That(events.Count(item => item.Kind == ExecutionNotificationKind.Issue && item.Phase == "testCleanup"), Is.EqualTo(2));
            Assert.That(events.Count(item => item.Kind == ExecutionNotificationKind.TestCompleted && item.Result.Outcome == TestOutcome.Succeeded), Is.EqualTo(2));
        }

        [Test]
        public void ModuleCleanupFailureStopsSubsequentModules()
        {
            var plan = Plan();
            plan.Modules.Add(new ExecutionModule { ProjectId = "project", Name = "Other", Tests = new List<ExecutionTest> { new ExecutionTest { Id = "other", Procedure = "Other" } } });
            var host = new Host { FailingHook = "ModuleAfter" };
            var events = Execute(plan, host);
            Assert.That(host.Trace, Does.Not.Contain("open:Other"));
            Assert.That(events.Single(item => item.Kind == ExecutionNotificationKind.Issue).Phase, Is.EqualTo("moduleCleanup"));
        }

        [Test]
        public void CancellationAtSafePointRunsCleanupAndReleasesGate()
        {
            var cancelled = false;
            var host = new Host();
            var events = Execute(Plan(), host, () => cancelled, () => cancelled = true);
            Assert.That(host.Trace, Contains.Item("After").And.Contains("ModuleAfter").And.Contains("stopFake"));
            Assert.That(host.Trace, Does.Not.Contain("One").And.Not.Contain("Two"));
            Assert.That(events.Last().Cancelled, Is.True);
            Assert.That(TestExecutionCore.IsRunning, Is.False);
        }

        [Test]
        public void NestedRunCannotCaptureTheActiveRunsAssertions()
        {
            var attempted = false;
            Execute(Plan(), new Host(), pump: () =>
            {
                attempted = true;
                Assert.Throws<InvalidOperationException>(() => Execute(Plan(), new Host()));
            });
            Assert.That(attempted, Is.True);
            Assert.That(TestExecutionCore.IsRunning, Is.False);
        }

        [Test]
        public void BodyComErrorProducesInconclusiveAndStillRunsCleanup()
        {
            var host = new Host { TestError = new COMException("body error") };
            var events = Execute(Plan(), host);
            Assert.That(events.Count(item => item.Kind == ExecutionNotificationKind.TestCompleted && item.Result.Outcome == TestOutcome.Inconclusive), Is.EqualTo(2));
            Assert.That(host.Trace.Count(item => item == "After"), Is.EqualTo(2));
        }

        [Test]
        public void InvalidPlanIsRejectedBeforeOpeningHostResources()
        {
            var plan = Plan();
            plan.Modules[0].Tests[1].Id = "one";
            var host = new Host();
            Assert.Throws<ArgumentException>(() => Execute(plan, host));
            Assert.That(host.Trace, Is.Empty);
        }

        [Test]
        public void ObserverFailureCannotLeaveExecutionGateLocked()
        {
            Assert.Throws<InvalidOperationException>(() => new TestExecutionCore().Execute(Plan(), new Host(), new Host(),
                () => false, () => { }, notification => { throw new InvalidOperationException("observer"); }));
            Assert.That(TestExecutionCore.IsRunning, Is.False);
        }

        [Test]
        public void PlanRoundTripRetainsMetadataAndDoesNotSerializeLegacyObjects()
        {
            var plan = Plan();
            plan.Modules[0].Tests[0].Tag = new object();
            var bytes = ExternalTestProtocol.Serialize(plan);
            var copy = ExternalTestProtocol.Deserialize<TestExecutionPlan>(bytes);
            copy.Validate();
            Assert.That(copy.Modules[0].Tests[0].Id, Is.EqualTo("one"));
            Assert.That(copy.Modules[0].Tests[0].Tag, Is.Null);
            Assert.That(Encoding.UTF8.GetString(bytes), Does.Not.Contain("Tag"));
        }

        [Test]
        public void SourceHashNormalizesOnlyLineEndings()
        {
            Assert.That(ExternalTestProtocol.SourceHash("abc"), Is.EqualTo("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"));
            Assert.That(ExternalTestProtocol.SourceHash("a\r\nb\r"), Is.EqualTo(ExternalTestProtocol.SourceHash("a\nb\n")));
            Assert.That(ExternalTestProtocol.SourceHash("a "), Is.Not.EqualTo(ExternalTestProtocol.SourceHash("a")));
        }

        [Test]
        public void SnapshotRejectsChangedProductionSourcesAndConditionalConstants()
        {
            var actual = new ExternalProjectSnapshot { ConditionalCompilation = "Feature=1", Sources = new List<ExternalSourceSnapshot> {
                new ExternalSourceSnapshot { Module = "Production", Sha256 = "abc" } } };
            var expected = ExternalTestProtocol.Deserialize<ExternalProjectSnapshot>(ExternalTestProtocol.Serialize(actual));
            ExternalTestExecutionService.ValidateSnapshot(expected, actual);
            expected.Sources[0].Sha256 = "stale";
            Assert.Throws<ArgumentException>(() => ExternalTestExecutionService.ValidateSnapshot(expected, actual));
            expected.Sources[0].Sha256 = "abc";
            expected.ConditionalCompilation = "Feature=0";
            Assert.Throws<ArgumentException>(() => ExternalTestExecutionService.ValidateSnapshot(expected, actual));
        }

        [Test]
        public void MissingRequiredPlanFieldsAreRejectedByProtocol()
        {
            Assert.Throws<SerializationException>(() => ExternalTestProtocol.Deserialize<TestExecutionPlan>(Encoding.UTF8.GetBytes("{}")));
        }
    }
}
