using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace Rubberduck.UnitTesting.Execution
{
    /// <summary>Parser-independent input to host execution. Names identify procedures in the live host.</summary>
    [DataContract]
    public sealed class TestExecutionPlan
    {
        [DataMember(Name = "modules", IsRequired = true)]
        public List<ExecutionModule> Modules { get; set; } = new List<ExecutionModule>();

        public void Validate()
        {
            if (Modules == null || Modules.Count == 0) throw new ArgumentException("A plan requires at least one module.");
            var modules = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var module in Modules)
            {
                if (module == null || string.IsNullOrWhiteSpace(module.ProjectId) || string.IsNullOrWhiteSpace(module.Name))
                    throw new ArgumentException("Each module requires a projectId and name.");
                if (!modules.Add(module.ProjectId + "\n" + module.Name)) throw new ArgumentException("Duplicate module.");
                if (module.Tests == null || module.Tests.Count == 0) throw new ArgumentException("Each module requires selected tests.");
                foreach (var hooks in new[] { module.ModuleInitialize, module.ModuleCleanup, module.TestInitialize, module.TestCleanup })
                {
                    if (hooks == null || hooks.Any(hook => hook == null || string.IsNullOrWhiteSpace(hook.Name)))
                        throw new ArgumentException("Invalid hook list.");
                }
                var procedures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var test in module.Tests)
                {
                    if (test == null || string.IsNullOrWhiteSpace(test.Id) || string.IsNullOrWhiteSpace(test.Procedure))
                        throw new ArgumentException("Each test requires an id and procedure.");
                    if (!ids.Add(test.Id)) throw new ArgumentException("Duplicate test id.");
                    if (!procedures.Add(test.Procedure)) throw new ArgumentException("Duplicate selected procedure.");
                }
            }
        }
    }

    [DataContract]
    public sealed class ExecutionModule
    {
        [DataMember(Name = "projectId", IsRequired = true)] public string ProjectId { get; set; }
        [DataMember(Name = "name", IsRequired = true)] public string Name { get; set; }
        [DataMember(Name = "moduleInitialize", IsRequired = true)] public List<ExecutionProcedure> ModuleInitialize { get; set; } = new List<ExecutionProcedure>();
        [DataMember(Name = "moduleCleanup", IsRequired = true)] public List<ExecutionProcedure> ModuleCleanup { get; set; } = new List<ExecutionProcedure>();
        [DataMember(Name = "testInitialize", IsRequired = true)] public List<ExecutionProcedure> TestInitialize { get; set; } = new List<ExecutionProcedure>();
        [DataMember(Name = "testCleanup", IsRequired = true)] public List<ExecutionProcedure> TestCleanup { get; set; } = new List<ExecutionProcedure>();
        [DataMember(Name = "tests", IsRequired = true)] public List<ExecutionTest> Tests { get; set; } = new List<ExecutionTest>();
        // Used only by the existing discovery adapter; never accepted across the external protocol.
        public object Tag { get; set; }
    }

    [DataContract]
    public sealed class ExecutionProcedure
    {
        [DataMember(Name = "name", IsRequired = true)] public string Name { get; set; }
        public object Tag { get; set; }
    }

    [DataContract]
    public sealed class ExecutionTest
    {
        [DataMember(Name = "id", IsRequired = true)] public string Id { get; set; }
        [DataMember(Name = "procedure", IsRequired = true)] public string Procedure { get; set; }
        [DataMember(Name = "category")] public string Category { get; set; }
        [DataMember(Name = "ignored")] public bool Ignored { get; set; }
        public object Tag { get; set; }
    }
}
