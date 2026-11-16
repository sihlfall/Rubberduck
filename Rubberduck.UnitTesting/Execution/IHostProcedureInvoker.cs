using System;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting.Execution
{
    /// <summary>Host invocation without a Rubberduck declaration or parser context.</summary>
    public interface IHostProcedureInvoker
    {
        void EnsureProjectReferencesUnitTesting(IVBProject project);
        void RunProcedure(ITypeLibWrapper typeLib, string module, string procedure);
        void RunTestProcedure(ITypeLibWrapper typeLib, string module, string procedure,
            EventHandler<AssertCompletedEventArgs> assertions, out long duration);
    }
}
