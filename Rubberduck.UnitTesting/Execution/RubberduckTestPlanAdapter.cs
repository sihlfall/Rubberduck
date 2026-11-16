using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

namespace Rubberduck.UnitTesting.Execution
{
    internal static class RubberduckTestPlanAdapter
    {
        public static TestExecutionPlan Create(IEnumerable<TestMethod> tests, RubberduckParserState state)
        {
            var plan = new TestExecutionPlan();
            foreach (var group in tests.GroupBy(test => test.Declaration.QualifiedModuleName))
            {
                var name = group.Key;
                plan.Modules.Add(new ExecutionModule
                {
                    ProjectId = name.ProjectId, Name = name.ComponentName,
                    ModuleInitialize = Hooks(TestDiscovery.FindModuleInitializeMethods(name, state)),
                    ModuleCleanup = Hooks(TestDiscovery.FindModuleCleanupMethods(name, state)),
                    TestInitialize = Hooks(TestDiscovery.FindTestInitializeMethods(name, state)),
                    TestCleanup = Hooks(TestDiscovery.FindTestCleanupMethods(name, state)),
                    Tests = group.Select(test => new ExecutionTest
                    {
                        Id = test.Declaration.QualifiedName.ToString(), Procedure = test.Declaration.QualifiedName.MemberName,
                        Category = test.Category.Name, Ignored = test.IsIgnored, Tag = test
                    }).ToList()
                });
            }
            return plan;
        }

        private static List<ExecutionProcedure> Hooks(IEnumerable<Declaration> declarations)
        {
            return declarations.Select(declaration => new ExecutionProcedure { Name = declaration.QualifiedName.MemberName, Tag = declaration }).ToList();
        }
    }

    internal sealed class DiscoveredTestExecutionHost : ITestExecutionHost
    {
        private readonly IVBEInteraction _interaction;
        private readonly ITypeLibWrapperProvider _wrappers;
        public DiscoveredTestExecutionHost(IVBEInteraction interaction, ITypeLibWrapperProvider wrappers)
        {
            _interaction = interaction;
            _wrappers = wrappers;
        }
        public ITestModuleSession OpenModule(ExecutionModule module) => new Session(_interaction, _wrappers.TypeLibWrapperFromProject(module.ProjectId));

        private sealed class Session : ITestModuleSession
        {
            private readonly IVBEInteraction _interaction;
            private readonly ITypeLibWrapper _wrapper;
            public Session(IVBEInteraction interaction, ITypeLibWrapper wrapper) { _interaction = interaction; _wrapper = wrapper; }
            public void RunHook(ExecutionProcedure procedure) => _interaction.RunDeclarations(_wrapper, new[] { (Declaration)procedure.Tag });
            public void RunTest(ExecutionTest test, EventHandler<AssertCompletedEventArgs> assertions, out long duration)
                => _interaction.RunTestMethod(_wrapper, (TestMethod)test.Tag, assertions, out duration);
            public void Dispose() => _wrapper.Dispose();
        }
    }
}
