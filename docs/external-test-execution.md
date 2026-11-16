# External test execution (full-package prototype)

Rubberduck's normal discovery/UI and externally supplied test plans now use the
same lifecycle/assertion execution core. The external route accepts procedure
names and hooks; it does not construct Rubberduck declarations or parse trees.
It does not publish per-test updates to the Rubberduck Test Explorer.

This milestone still requires the full add-in. The parser remains installed and
can run normally. Its suspension lock prevents background parsing from racing
host execution, but an external plan does not require Rubberduck's parser to
reach Ready or supply discovery results. Standalone packaging and VS Code
engine profiles are separate work.

## Enable and discover

Set `RUBBERDUCK_TEST_PIPE=1` in the environment of the host process before the
add-in starts. The endpoint is off by default. It starts after full application
startup and is closed before VBE services are torn down.

Each opted-in host listens on `Rubberduck.Tests.<host-process-id>`. The pipe ACL
grants access only to the current Windows user. No TCP listener or additional
COM registration is introduced. Choose the process explicitly; the bridge does
not start hosts, attach to other processes, import code, or modify source.

From Windows PowerShell:

```powershell
development/Invoke-ExternalTests.ps1 -HostProcessId 1234 -Command describe
development/Invoke-ExternalTests.ps1 -HostProcessId 1234 -Command run -RequestFile plan.json
# From another console while that run's stream remains open:
development/Invoke-ExternalTests.ps1 -HostProcessId 1234 -Command cancel -RunId example-1
```

The description supplies host kind, bitness, and the live project IDs,
conditional-compilation arguments, and source fingerprints. Protected projects
are omitted. Project IDs are host-session identities; do not substitute IDs from
the external parser. Host/module/procedure names bind the external metadata to
the actual loaded project.

## Version 1 request

```json
{
  "version": 1,
  "command": "run",
  "runId": "example-1",
  "hostKind": "<hostKind from describe>",
  "bitness": 64,
  "timeoutSeconds": 60,
  "projects": [{
    "projectId": "<live project ID>",
    "conditionalCompilation": "",
    "sources": [
      { "module": "ArithmeticTests", "sha256": "<hash of the source that was parsed>" }
    ]
  }],
  "plan": {
    "modules": [{
      "projectId": "<live project ID>",
      "name": "ArithmeticTests",
      "moduleInitialize": [{ "name": "StartModule" }],
      "moduleCleanup": [{ "name": "FinishModule" }],
      "testInitialize": [{ "name": "BeforeEach" }],
      "testCleanup": [{ "name": "AfterEach" }],
      "tests": [{
        "id": "client-stable-addition-id",
        "procedure": "Addition",
        "category": "Arithmetic",
        "ignored": false
      }]
    }]
  }
}
```

All four hook arrays are required, including empty arrays. Selected tests are
explicit; duplicate IDs/procedures/modules are rejected. Tests and hooks must
identify parameterless Subs in standard modules. The client owns discovery,
annotation validation, categories, and source navigation. Private procedures
are invoked through the existing host type-library implementation.

For every selected project, include fingerprints for **all code-bearing
components**, including production modules, classes, forms, and document modules.
The fingerprint is SHA256 of UTF-8 code-pane text, without a BOM, with CRLF and CR
normalized to LF. No whitespace is trimmed. Export-only Attribute/designer records
are not code-pane text. Compute fingerprints from the actual parsed snapshot;
copying fresh host fingerprints onto stale parsed data defeats validation.

The bridge checks host kind/bitness, the complete component set, all hashes, and
the raw conditional-compilation string before reference setup and host compilation.
The external frontend must also use matching built-in constants, user defines,
and references. This prototype does not import, rewrite, or synchronize source,
nor does it replace the host compiler.

## Transport and results

One request per duplex named-pipe connection. Frames consist of a four-byte
little-endian positive byte length followed by UTF-8 JSON; maximum length is
1 MiB in either direction. A describe/cancel command returns one frame. A run
streams `run_started`, `test_started`, `test_completed`, `issue`, and
`run_completed`. Every event carries version 1; run events retain runId and test
events retain client IDs. Durations are milliseconds. Test statuses preserve
`Succeeded`, `Failed`, `Ignored`, `Inconclusive`, and `Unknown`.

An `error` frame indicates request/preflight/execution-control failure. Transport
EOF before a terminal frame is incomplete execution. Module/test-hook problems
also emit separate `issue` events, so clients must not infer overall success
solely from body results. The client script outputs JSON frames and fails on
failed/inconclusive/unknown tests, issues, cancellation, or protocol errors.

Execution intentionally preserves the full runner's policy: first nonsuccessful
body assertion determines its result; hook assertions do not become body results;
failed module initialization blocks its bodies; failed test initialization skips
that body and its cleanup; test-cleanup COM errors allow subsequent tests; module
cleanup COM errors abort subsequent modules. Results are published before cleanup.
The existing UI keeps its established result behavior; external clients additionally
receive hook issues.

Only one core execution may run at a time in a host because assertion events and
VBA globals are shared. The external service also rejects overlapping pending runs.
Cancellation is cooperative between procedures; it cannot interrupt an infinite
VBA call, guarantee cleanup after a host reset, or dismiss host error dialogs.
The timeout requests that same cooperative cancellation. Disconnecting the run
client requests cancellation. A separate cancel connection targets the exact runId.
Slow clients cannot block the host UI thread: a bounded event queue cancels a run
if it fills. Shutdown closes the pipes without waiting on the host UI thread.

## Code boundaries

- `Rubberduck.UnitTesting/Execution/TestExecutionCore.cs`: shared execution policy.
- `RubberduckTestPlanAdapter.cs`: existing discovery to plans and the legacy
  procedure-invocation adapter, preserving `ITestEngine`/`TestMethod` consumers.
- `IHostProcedureInvoker`: names-only host invocation, implemented alongside
  `IVBEInteraction` by `VBEInteraction`.
- `ExternalTestExecutionService`: full-package host coordination and preflight.
- `Rubberduck.Main/ExternalTesting/ExternalTestPipeServer.cs`: transport.
- `RubberduckIoCInstaller` and `Extension`: registration and opt-in lifecycle.

Assertion COM classes, GUIDs, type-library generation, fakes, deployment, the test
grid performance fix, and parser implementation stay in their existing package.
The core is logically independent of discovery; extracting an independently
distributable assembly remains part of the runner-only milestone.

## Validation for this milestone

The unit-testing assembly, full add-in, and test assembly were compiled for
.NET Framework 4.6.2 using VS 2022 MSBuild and workspace-local NuGet reference
assemblies. Validation used an isolated tracked-source copy, excluding pre-existing
untracked UI work, with the legacy analyzer/deployment bootstrap disabled only
in that validation copy.

The 187 unit-testing regression cases passed, followed by two added service
integration cases. These cover the old engine/UI behavior, shared lifecycle
policy, snapshots, parser-to-plan adaptation, and real named-pipe framing,
cancellation, disconnects, and preflight error delivery. The service cases use
a mock VBE with the real parser-suspension coordinator: an externally supplied
plan executes without any Rubberduck parse or test annotations, and a stale
snapshot fails before compilation while releasing its run reservation.

Execution inside an installed Excel/VB6 host has not yet been exercised for this
branch. No installed add-in or COM registration was changed by validation.
