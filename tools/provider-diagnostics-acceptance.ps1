[CmdletBinding()]
param(
    [string]$CorePath = '.\src\OMNIX.Core\bin\Release\OMNIX.Core.dll',
    [string]$OutputPath = '.\build\artifact\provider-diagnostics-acceptance.json'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$core = (Resolve-Path -LiteralPath $CorePath).Path
$dependency = Join-Path (Split-Path -Parent $core) 'Newtonsoft.Json.dll'
if (Test-Path $dependency) { [void][Reflection.Assembly]::LoadFrom($dependency) }
[void][Reflection.Assembly]::LoadFrom($core)

# Exercises the compiled production diagnostic path with deterministic adapters.
# No cloud request, API key or Office application is involved.
$source = @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.AiGateway;
using OMNIX.Core.AiGateway.Adapters;
using OMNIX.Core.Errors;
using OMNIX.Core.Settings;

public sealed class DiagnosticAdapter : IProviderAdapter
{
    public ProviderInfo Info { get; private set; }
    public int Sends;
    public int CatalogReads;
    public bool Empty;
    public bool FailAuth;
    public ChatRequest Request;
    public ProviderCredentials Credentials;
    public DiagnosticAdapter(ProviderKind kind)
    { Info = new ProviderInfo { Id = "diagnostic-fixture", DisplayName = "Diagnostic fixture", Kind = kind }; }
    public void Configure(ProviderCredentials credentials) { Credentials = credentials; }
    public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct)
    { CatalogReads++; throw new InvalidOperationException("This server has no model catalog."); }
    public Task<bool> TestConnectionAsync(CancellationToken ct)
    { throw new InvalidOperationException("Legacy catalog test must not run."); }
    public bool SupportsVisionNow() { return false; }
    public Task<ChatResponse> SendAsync(ChatRequest request, Action<string> delta, CancellationToken ct)
    {
        Sends++; Request = request;
        ct.ThrowIfCancellationRequested();
        if (FailAuth) throw OmnixException.Auth("Synthetic authentication failure.");
        return Task.FromResult(new ChatResponse { Text = Empty ? "   " : "OK" });
    }
}

public static class ProviderDiagnosticsHarness
{
    private static void Probe(DiagnosticAdapter adapter, PrivacyGate gate, CancellationToken ct)
    { ProviderDiagnostics.TestModelAsync(adapter, new ProviderCredentials { Model = "private/manual-model" }, gate, ct).GetAwaiter().GetResult(); }

    public static Dictionary<string, bool> Run()
    {
        var checks = new Dictionary<string, bool>();
        var settings = SettingsManager.Instance.Settings;
        var oldPrivacy = settings.Privacy;
        try
        {
            settings.Privacy = PrivacyMode.CloudAllowed;
            var gate = new PrivacyGate();
            var adapter = new DiagnosticAdapter(ProviderKind.Cloud);
            Probe(adapter, gate, CancellationToken.None);
            checks["ManualModelWorksWithoutCatalog"] = adapter.Sends == 1 && adapter.CatalogReads == 0 &&
                adapter.Credentials.Model == "private/manual-model";
            checks["ProbeContainsOnlySyntheticText"] = adapter.Request.History == null && adapter.Request.SystemPrompt == null &&
                adapter.Request.UserTurn.Text == "Reply with OK." && !adapter.Request.UserTurn.HasImages;

            adapter = new DiagnosticAdapter(ProviderKind.Cloud) { Empty = true };
            try { Probe(adapter, gate, CancellationToken.None); checks["EmptyResponseFails"] = false; }
            catch (OmnixException ex) { checks["EmptyResponseFails"] = ex.Code == ErrorCode.PROVIDER_ERROR; }

            adapter = new DiagnosticAdapter(ProviderKind.Cloud) { FailAuth = true };
            try { Probe(adapter, gate, CancellationToken.None); checks["AuthFailurePreserved"] = false; }
            catch (OmnixException ex) { checks["AuthFailurePreserved"] = ex.Code == ErrorCode.AUTH_ERROR; }

            settings.Privacy = PrivacyMode.LocalOnly;
            adapter = new DiagnosticAdapter(ProviderKind.Cloud);
            try { Probe(adapter, gate, CancellationToken.None); checks["LocalOnlyBlocksCloudBeforeSend"] = false; }
            catch (OmnixException ex) { checks["LocalOnlyBlocksCloudBeforeSend"] = ex.Code == ErrorCode.PRIVACY_BLOCKED && adapter.Sends == 0; }

            adapter = new DiagnosticAdapter(ProviderKind.Local);
            Probe(adapter, gate, CancellationToken.None);
            checks["LocalOnlyAllowsLocalTest"] = adapter.Sends == 1;

            settings.Privacy = PrivacyMode.AskBeforeSending;
            adapter = new DiagnosticAdapter(ProviderKind.Cloud);
            gate.CloudConfirmationCallback = name => Task.FromResult(Tuple.Create(false, false));
            try { Probe(adapter, gate, CancellationToken.None); checks["DeclinedConsentBlocksSend"] = false; }
            catch (OmnixException ex) { checks["DeclinedConsentBlocksSend"] = ex.Code == ErrorCode.PRIVACY_BLOCKED && adapter.Sends == 0; }

            adapter = new DiagnosticAdapter(ProviderKind.Cloud);
            ProviderDiagnostics.TestSyntheticModelAsync(adapter, new ProviderCredentials { Model="manual-model" }, CancellationToken.None).GetAwaiter().GetResult();
            checks["ExplicitDiagnosticDoesNotAskDocumentConsent"] = adapter.Sends == 1 && adapter.Request.History == null && adapter.Request.SystemPrompt == null && !adapter.Request.HasImages && adapter.Request.UserTurn.Text == "Reply with OK.";
            settings.Privacy = PrivacyMode.LocalOnly;
            adapter = new DiagnosticAdapter(ProviderKind.Cloud);
            try { ProviderDiagnostics.TestSyntheticModelAsync(adapter, new ProviderCredentials(), CancellationToken.None).GetAwaiter().GetResult(); checks["ExplicitDiagnosticRespectsLocalOnly"]=false; }
            catch(OmnixException ex) { checks["ExplicitDiagnosticRespectsLocalOnly"]=ex.Code==ErrorCode.PRIVACY_BLOCKED && adapter.Sends==0; }
            settings.Privacy = PrivacyMode.AskBeforeSending;

            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                adapter = new DiagnosticAdapter(ProviderKind.Local);
                try { Probe(adapter, gate, cts.Token); checks["CancellationBlocksSend"] = false; }
                catch (OperationCanceledException) { checks["CancellationBlocksSend"] = adapter.Sends == 0; }
            }

            var models = Enumerable.Range(0, 750).Select(i => "model-" + i).ToList();
            var options = ProviderDiagnostics.ModelOptions(models, "private/manual-model");
            checks["CatalogBeyond300AndManualModelKept"] = options.Contains("model-749") && options.Contains("private/manual-model") && options.Count == 751;
            checks["EmptyCatalogKeepsManualModel"] = ProviderDiagnostics.ModelOptions(null, "manual").SequenceEqual(new[] { "manual" });
            checks["DuplicateModelsRemoved"] = ProviderDiagnostics.ModelOptions(new[] { "a", "a", "", null, "b" }, "a").SequenceEqual(new[] { "a", "b" });

            checks["CompletionUrlNormalized"] = CustomOpenAiCompatibleAdapter.NormalizeBaseUrl("https://example.invalid/gateway/v1/chat/completions/") == "https://example.invalid/gateway/v1";
            checks["ModelUrlNormalized"] = CustomOpenAiCompatibleAdapter.NormalizeBaseUrl("https://example.invalid/api/v1/models") == "https://example.invalid/api/v1";
            checks["LocalBaseUrlPreserved"] = CustomOpenAiCompatibleAdapter.NormalizeBaseUrl("http://127.0.0.1:1234/v1/") == "http://127.0.0.1:1234/v1";
            foreach (string bad in new[] { "http://example.invalid/v1", "https://user:password@example.invalid/v1", "https://example.invalid/v1?key=synthetic", "https://example.invalid/v1#fragment" })
            {
                try { CustomOpenAiCompatibleAdapter.NormalizeBaseUrl(bad); checks["InvalidEndpointRejected-" + checks.Count] = false; }
                catch (OmnixException) { checks["InvalidEndpointRejected-" + checks.Count] = true; }
            }
        }
        finally { settings.Privacy = oldPrivacy; }
        return checks;
    }
}
'@
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @($core, 'System.Core.dll')
$checks = [ProviderDiagnosticsHarness]::Run()
$failures = @($checks.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
$report = [ordered]@{
    TestId = 'PROVIDER-DIAGNOSTICS-RUNTIME-001'
    SourceCommit = (git rev-parse HEAD).Trim()
    GeneratedUtc = [DateTime]::UtcNow.ToString('o')
    Checks = $checks
    CheckCount = $checks.Count
    FailureCount = $failures.Count
    Failures = $failures
    OverallPass = ($failures.Count -eq 0)
}
$outDir = Split-Path -Parent $OutputPath
if ($outDir) { New-Item -ItemType Directory -Force $outDir | Out-Null }
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$report | ConvertTo-Json -Depth 8
if ($failures.Count -gt 0) { exit 1 }
