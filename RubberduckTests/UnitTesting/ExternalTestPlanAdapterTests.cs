using System.Linq;
using NUnit.Framework;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.Execution;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    public class ExternalTestPlanAdapterTests
    {
        [Test]
        public void ExistingDiscoveryProducesTheSameSelectedTestsAndLifecyclePlan()
        {
            const string code = @"Option Explicit
Option Private Module

'@TestModule

Private Assert As Object

'@ModuleInitialize
Private Sub StartModule()
End Sub
'@ModuleCleanup
Private Sub EndModule()
End Sub
'@TestInitialize
Private Sub BeforeEach()
End Sub
'@TestCleanup
Private Sub AfterEach()
End Sub
'@TestMethod(""Numbers"")
Private Sub First()
End Sub
'@TestMethod
'@IgnoreTest
Private Sub Second()
End Sub";
            // Model the VBE's CRLF code-pane text independently of checkout line endings.
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code.Replace("\r\n", "\n").Replace("\n", "\r\n"), out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                Assert.That(state.Status, Is.EqualTo(ParserState.Ready),
                    string.Join("; ", state.ModuleExceptions.Select(error => error.Item2.Message)));
                var tests = TestDiscovery.GetAllTests(state).OrderBy(test => test.Declaration.IdentifierName).ToList();
                Assert.That(tests.Count, Is.EqualTo(2),
                    string.Join("; ", state.AllUserDeclarations.Select(declaration => declaration.IdentifierName + ":" + declaration.DeclarationType)));
                var plan = RubberduckTestPlanAdapter.Create(tests, state);
                plan.Validate();
                var module = plan.Modules.Single();
                CollectionAssert.AreEqual(new[] { "First", "Second" }, module.Tests.Select(test => test.Procedure));
                Assert.That(module.Tests[0].Category, Is.EqualTo("Numbers"));
                Assert.That(module.Tests[1].Ignored, Is.True);
                Assert.That(module.Tests[0].Tag, Is.SameAs(tests[0]));
                Assert.That(module.ModuleInitialize.Single().Name, Is.EqualTo("StartModule"));
                Assert.That(module.ModuleCleanup.Single().Name, Is.EqualTo("EndModule"));
                Assert.That(module.TestInitialize.Single().Name, Is.EqualTo("BeforeEach"));
                Assert.That(module.TestCleanup.Single().Name, Is.EqualTo("AfterEach"));
            }
        }
    }
}
