using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.UnitTesting.Execution
{
    [DataContract]
    public sealed class ExternalTestRequest
    {
        [DataMember(Name = "version", IsRequired = true)] public int Version { get; set; }
        [DataMember(Name = "command", IsRequired = true)] public string Command { get; set; }
        [DataMember(Name = "runId")] public string RunId { get; set; }
        [DataMember(Name = "hostKind")] public string HostKind { get; set; }
        [DataMember(Name = "bitness")] public int Bitness { get; set; }
        [DataMember(Name = "timeoutSeconds")] public int TimeoutSeconds { get; set; }
        [DataMember(Name = "projects")] public List<ExternalProjectSnapshot> Projects { get; set; }
        [DataMember(Name = "plan")] public TestExecutionPlan Plan { get; set; }
    }

    [DataContract]
    public sealed class ExternalProjectSnapshot
    {
        [DataMember(Name = "projectId", IsRequired = true)] public string ProjectId { get; set; }
        [DataMember(Name = "name")] public string Name { get; set; }
        [DataMember(Name = "path")] public string Path { get; set; }
        [DataMember(Name = "conditionalCompilation", IsRequired = true)] public string ConditionalCompilation { get; set; }
        [DataMember(Name = "sources", IsRequired = true)] public List<ExternalSourceSnapshot> Sources { get; set; }
    }

    [DataContract]
    public sealed class ExternalSourceSnapshot
    {
        [DataMember(Name = "module", IsRequired = true)] public string Module { get; set; }
        [DataMember(Name = "sha256", IsRequired = true)] public string Sha256 { get; set; }
    }

    [DataContract]
    public sealed class ExternalTestEvent
    {
        [DataMember(Name = "version")] public int Version { get; set; } = 1;
        [DataMember(Name = "event")] public string Event { get; set; }
        [DataMember(Name = "runId", EmitDefaultValue = false)] public string RunId { get; set; }
        [DataMember(Name = "testId", EmitDefaultValue = false)] public string TestId { get; set; }
        [DataMember(Name = "module", EmitDefaultValue = false)] public string Module { get; set; }
        [DataMember(Name = "status", EmitDefaultValue = false)] public string Status { get; set; }
        [DataMember(Name = "phase", EmitDefaultValue = false)] public string Phase { get; set; }
        [DataMember(Name = "message", EmitDefaultValue = false)] public string Message { get; set; }
        [DataMember(Name = "durationMs")] public long DurationMs { get; set; }
        [DataMember(Name = "cancelled")] public bool Cancelled { get; set; }
        [DataMember(Name = "hostKind", EmitDefaultValue = false)] public string HostKind { get; set; }
        [DataMember(Name = "bitness", EmitDefaultValue = false)] public int Bitness { get; set; }
        [DataMember(Name = "projects", EmitDefaultValue = false)] public List<ExternalProjectSnapshot> Projects { get; set; }
    }

    public interface IExternalTestExecutionService
    {
        Task<ExternalTestEvent> DescribeAsync(CancellationToken token);
        Task RunAsync(ExternalTestRequest request, Action<ExternalTestEvent> publish, CancellationToken token);
        bool Cancel(string runId);
    }

    public static class ExternalTestProtocol
    {
        public const int Version = 1;
        public const int MaximumFrameBytes = 1024 * 1024;

        public static byte[] Serialize<T>(T value)
        {
            using (var stream = new MemoryStream())
            {
                new DataContractJsonSerializer(typeof(T)).WriteObject(stream, value);
                return stream.ToArray();
            }
        }

        public static T Deserialize<T>(byte[] bytes)
        {
            if (bytes == null || bytes.Length > MaximumFrameBytes) throw new InvalidDataException("Invalid protocol frame size.");
            using (var stream = new MemoryStream(bytes))
                return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(stream);
        }

        /// <summary>SHA256 of UTF-8 code-pane text, with CRLF/CR normalized to LF; no trimming or BOM.</summary>
        public static string SourceHash(string code)
        {
            var normalized = (code ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(new UTF8Encoding(false).GetBytes(normalized))).Replace("-", "").ToLowerInvariant();
        }
    }
}
