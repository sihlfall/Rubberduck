using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.Execution;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    [TestFixture, NonParallelizable]
    public class ExternalTestExecutionServiceTests
    {
        private delegate void InvokeTest(ITypeLibWrapper wrapper, string module, string procedure,
            EventHandler<AssertCompletedEventArgs> assertions, out long duration);

        private sealed class Fixture : IDisposable
        {
            public readonly RubberduckParserState State;
            public readonly ExternalTestExecutionService Service;
            public readonly Mock<IHostProcedureInvoker> Invoker = new Mock<IHostProcedureInvoker>();
            public readonly Mock<IVBETypeLibsAPI> TypeLibs = new Mock<IVBETypeLibsAPI>();
            public readonly Mock<IUiDispatcher> Dispatcher = new Mock<IUiDispatcher>();
            public readonly List<string> Invocations = new List<string>();

            public Fixture()
            {
                // Deliberately no Rubberduck test annotations and no parse request.
                var vbe = MockVbeBuilder.BuildFromSingleStandardModule("Private Sub ClientSelected()\r\nEnd Sub\r\n", out _);
                State = MockParser.Create(vbe.Object).State;
                Dispatcher.Setup(dispatcher => dispatcher.StartTask(It.IsAny<Action>(), It.IsAny<CancellationToken>(), It.IsAny<TaskCreationOptions>()))
                    .Returns((Action action, CancellationToken token, TaskCreationOptions options) =>
                    {
                        token.ThrowIfCancellationRequested(); action(); return Task.CompletedTask;
                    });
                Dispatcher.Setup(dispatcher => dispatcher.StartTask(It.IsAny<Func<ExternalTestEvent>>(), It.IsAny<CancellationToken>(), It.IsAny<TaskCreationOptions>()))
                    .Returns((Func<ExternalTestEvent> action, CancellationToken token, TaskCreationOptions options) => Task.FromResult(action()));
                var wrappers = new Mock<ITypeLibWrapperProvider>();
                wrappers.Setup(provider => provider.TypeLibWrapperFromProject(It.IsAny<string>())).Returns(new Mock<ITypeLibWrapper>().Object);
                TypeLibs.Setup(api => api.GetProjectConditionalCompilationArgsRaw(It.IsAny<IVBProject>())).Returns("Feature=1");
                TypeLibs.Setup(api => api.CompileProject(It.IsAny<IVBProject>())).Returns(true);
                var fakes = new Mock<IFakesFactory>();
                fakes.Setup(factory => factory.Create()).Returns(new Mock<IFakes>().Object);
                long ignored;
                Invoker.Setup(invoker => invoker.RunTestProcedure(It.IsAny<ITypeLibWrapper>(), It.IsAny<string>(), It.IsAny<string>(),
                        It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out ignored))
                    .Callback(new InvokeTest((ITypeLibWrapper wrapper, string module, string procedure,
                        EventHandler<AssertCompletedEventArgs> assertions, out long duration) =>
                    {
                        duration = 12;
                        Invocations.Add(module + "." + procedure);
                        assertions(this, new AssertCompletedEventArgs(TestOutcome.Succeeded));
                    }));
                Service = new ExternalTestExecutionService(State, Dispatcher.Object, vbe.Object,
                    (IProjectsRepository)State.ProjectsProvider, wrappers.Object, TypeLibs.Object, Invoker.Object, fakes.Object);
            }

            public async Task<ExternalTestRequest> Request()
            {
                var description = await Service.DescribeAsync(CancellationToken.None);
                var project = description.Projects.Single();
                return new ExternalTestRequest
                {
                    Version = 1, Command = "run", RunId = Guid.NewGuid().ToString("N"), TimeoutSeconds = 30,
                    HostKind = description.HostKind, Bitness = description.Bitness, Projects = description.Projects,
                    Plan = new TestExecutionPlan { Modules = new List<ExecutionModule> {
                        new ExecutionModule {
                            ProjectId = project.ProjectId, Name = project.Sources.Single().Module,
                            Tests = new List<ExecutionTest> { new ExecutionTest { Id = "from-external-parser", Procedure = "ClientSelected" } }
                        }
                    } }
                };
            }
            public void Dispose() => State.Dispose();
        }

        [Test, Timeout(15000)]
        public async Task ExternalPlanExecutesWithoutRubberduckParsingOrAnnotations()
        {
            using (var fixture = new Fixture())
            {
                var request = await fixture.Request();
                Assert.That(fixture.State.Status, Is.Not.EqualTo(ParserState.Ready));
                Assert.That(TestDiscovery.GetAllTests(fixture.State), Is.Empty);
                var events = new List<ExternalTestEvent>();
                await fixture.Service.RunAsync(request, events.Add, CancellationToken.None);
                Assert.That(fixture.Invocations.Single(), Does.EndWith(".ClientSelected"));
                var result = events.Single(item => item.Event == "test_completed");
                Assert.That(result.TestId, Is.EqualTo("from-external-parser"));
                Assert.That(result.Status, Is.EqualTo("Succeeded"));
                Assert.That(result.DurationMs, Is.EqualTo(12));
                Assert.That(events.Last().Event, Is.EqualTo("run_completed"));
            }
        }

        [Test, Timeout(15000)]
        public async Task StaleSourceIsRejectedBeforeCompilationAndReservationCanBeReused()
        {
            using (var fixture = new Fixture())
            {
                var stale = await fixture.Request();
                stale.Projects[0].Sources[0].Sha256 = "stale";
                Assert.ThrowsAsync<InvalidOperationException>(async () => await fixture.Service.RunAsync(stale, item => { }, CancellationToken.None));
                fixture.TypeLibs.Verify(api => api.CompileProject(It.IsAny<IVBProject>()), Times.Never);
                Assert.That(fixture.Invocations, Is.Empty);
                var fresh = await fixture.Request();
                await fixture.Service.RunAsync(fresh, item => { }, CancellationToken.None);
                Assert.That(fixture.Invocations.Count, Is.EqualTo(1));
            }
        }
    }
}
