# OMNIX runtime diagnostic journal acceptance
#
# Deterministic, offline test of the two new runtime logs. It proves correlation/timing events are
# emitted and that secret-like metadata is redacted. No provider call and no Office document is used.

[CmdletBinding()]
param(
    [string]$CorePath = ".\src\OMNIX.Core\bin\Release\OMNIX.Core.dll",
    [string]$OutputPath = ".\build\artifact\runtime-diagnostic-journal-acceptance.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$core = (Resolve-Path -LiteralPath $CorePath -ErrorAction Stop).Path
$coreDir = Split-Path -Parent $core
$newtonsoft = Join-Path $coreDir 'Newtonsoft.Json.dll'
if (Test-Path -LiteralPath $newtonsoft) { [void][Reflection.Assembly]::LoadFrom($newtonsoft) }
[void][Reflection.Assembly]::LoadFrom($core)

$source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json.Linq;
using OMNIX.Core.Logging;

public sealed class RuntimeDiagnosticJournalResult
{
    public string TestId { get; set; }
    public int EvidenceSchema { get; set; }
    public string GeneratedUtc { get; set; }
    public bool HumanLogCreatedPass { get; set; }
    public bool StructuredLogCreatedPass { get; set; }
    public bool CorrelationPass { get; set; }
    public bool ToolTimingPass { get; set; }
    public bool SecretRedactionPass { get; set; }
    public bool NoPayloadContentPass { get; set; }
    public int FailureCount { get; set; }
    public List<string> Failures { get; set; }
    public bool OverallPass { get; set; }
    public string Privacy { get; set; }
}

public static class RuntimeDiagnosticJournalHarness
{
    public static RuntimeDiagnosticJournalResult Run()
    {
        var failures = new List<string>();
        var result = new RuntimeDiagnosticJournalResult
        {
            TestId = "RUNTIME-DIAGNOSTIC-JOURNAL-001",
            EvidenceSchema = 1,
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            Failures = failures,
            Privacy = "Evidence contains booleans/counts only. Injected secret markers, prompts, Office content and tool arguments are not emitted."
        };

        string temp = Path.Combine(Path.GetTempPath(), "omnix-journal-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);

        var field = typeof(Logger).GetField("_baseDir", BindingFlags.Static | BindingFlags.NonPublic);
        if (field == null) throw new InvalidOperationException("Logger._baseDir field not found.");
        field.SetValue(null, temp);

        string marker = "OMNIX_SECRET_" + Guid.NewGuid().ToString("N");
        string fakePrompt = "PROMPT_CONTENT_" + Guid.NewGuid().ToString("N");
        string fakeDocument = "DOCUMENT_CONTENT_" + Guid.NewGuid().ToString("N");

        RuntimeDiagnosticJournal.BeginRequest("Excel", true, fakePrompt.Length, 3, false);
        RuntimeDiagnosticJournal.SetProvider("custom", "test-model");
        RuntimeDiagnosticJournal.Event("tool_execute_start", "write_to_cell", "write", null, null,
            "api_key=" + marker + "; baseUrl=https://private.invalid/v1; safe=metadata");
        RuntimeDiagnosticJournal.Event("tool_execute_end", "write_to_cell", "success", 123, null,
            "successfulWrites=1; failedWrites=0");
        RuntimeDiagnosticJournal.CompleteRequest("success", 1, 0, true);

        string journeyPath = Path.Combine(temp, "logs", "runtime-journey.log");
        string jsonlPath = Path.Combine(temp, "logs", "runtime-events.jsonl");
        result.HumanLogCreatedPass = File.Exists(journeyPath) && new FileInfo(journeyPath).Length > 0;
        result.StructuredLogCreatedPass = File.Exists(jsonlPath) && new FileInfo(jsonlPath).Length > 0;
        if (!result.HumanLogCreatedPass) failures.Add("Human runtime journey log was not created.");
        if (!result.StructuredLogCreatedPass) failures.Add("Structured runtime JSONL log was not created.");

        string human = File.Exists(journeyPath) ? File.ReadAllText(journeyPath) : "";
        string raw = File.Exists(jsonlPath) ? File.ReadAllText(jsonlPath) : "";
        var events = new List<JObject>();
        if (File.Exists(jsonlPath))
        {
            foreach (string line in File.ReadAllLines(jsonlPath))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                events.Add(JObject.Parse(line));
            }
        }

        var traceIds = events.Select(e => (string)e["traceId"]).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToList();
        result.CorrelationPass = events.Count >= 5 && traceIds.Count == 1 &&
            events.Any(e => (string)e["event"] == "request_start") &&
            events.Any(e => (string)e["event"] == "request_complete");
        if (!result.CorrelationPass) failures.Add("Structured events did not preserve one request correlation id.");

        result.ToolTimingPass = events.Any(e =>
            (string)e["event"] == "tool_execute_end" &&
            (string)e["tool"] == "write_to_cell" &&
            (long?)e["elapsedMs"] == 123 &&
            (string)e["status"] == "success");
        if (!result.ToolTimingPass) failures.Add("Tool timing/status event was not emitted.");

        result.SecretRedactionPass =
            human.IndexOf(marker, StringComparison.Ordinal) < 0 &&
            raw.IndexOf(marker, StringComparison.Ordinal) < 0 &&
            human.IndexOf("private.invalid", StringComparison.OrdinalIgnoreCase) < 0 &&
            raw.IndexOf("private.invalid", StringComparison.OrdinalIgnoreCase) < 0;
        if (!result.SecretRedactionPass) failures.Add("Secret-like detail or private Base URL leaked into runtime diagnostics.");

        result.NoPayloadContentPass =
            human.IndexOf(fakePrompt, StringComparison.Ordinal) < 0 &&
            raw.IndexOf(fakePrompt, StringComparison.Ordinal) < 0 &&
            human.IndexOf(fakeDocument, StringComparison.Ordinal) < 0 &&
            raw.IndexOf(fakeDocument, StringComparison.Ordinal) < 0 &&
            raw.IndexOf("ArgumentsJson", StringComparison.OrdinalIgnoreCase) < 0;
        if (!result.NoPayloadContentPass) failures.Add("Prompt/document/tool-argument content leaked into runtime diagnostics.");

        result.FailureCount = failures.Count;
        result.OverallPass = failures.Count == 0;
        try { Directory.Delete(temp, true); } catch { }
        return result;
    }
}
'@

Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @($core, $newtonsoft) -ErrorAction Stop
$result = [RuntimeDiagnosticJournalHarness]::Run()

$outDir = Split-Path -Parent $OutputPath
if ($outDir) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$result | ConvertTo-Json -Depth 8

if (-not $result.OverallPass) { exit 1 }
exit 0
