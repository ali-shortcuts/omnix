# OMNIX provider diagnostic/privacy anti-drift contract.
# Structural only. Compiled marker-redaction behavior is executed by request-budget-runtime CI.
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

function Read-Repo([string]$relative) {
    $path = Join-Path $root $relative
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $failures.Add("Missing: $relative")
        return ''
    }
    return Get-Content -LiteralPath $path -Raw
}
function Require([string]$relative,[string]$needle,[string]$reason) {
    $text = Read-Repo $relative
    if (-not $text.Contains($needle)) { $failures.Add("${relative}: missing '$needle' — $reason") }
}
function Forbid([string]$relative,[string]$needle,[string]$reason) {
    $text = Read-Repo $relative
    if ($text.Contains($needle)) { $failures.Add("${relative}: forbidden '$needle' — $reason") }
}
function Parse-Ps([string]$relative) {
    $path = Join-Path $root $relative
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $failures.Add("Missing: $relative"); return }
    $tokens=$null; $errors=$null
    [void][System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
    foreach($e in @($errors)){ $failures.Add("${relative}: parser error — $($e.Message)") }
}

$contracts = 'src/OMNIX.Core/AiGateway/ProviderContracts.cs'
Require $contracts 'provider response body' 'provider bodies must be treated as untrusted/private diagnostic input.'
Require $contracts 'provider_response_body=REDACTED' 'mapped provider diagnostics must record an explicit redaction marker.'
Require $contracts 'data_policy_no_eligible_endpoint' 'OpenRouter privacy/data-policy failures need a stable sanitized category.'
Require $contracts 'settings=https://openrouter.ai/settings/privacy' 'OpenRouter privacy remediation may expose only the official settings route, not the raw provider body.'
Require $contracts 'SafeProviderName' 'provider labels inserted into diagnostics must be sanitized.'
Forbid $contracts 'Body: " + trimmed' 'raw provider bodies must never be concatenated into mapped diagnostics.'
Forbid $contracts 'Body: " + detectionBody' 'detection-only provider bodies must never be emitted.'

$errors = 'src/OMNIX.Core/Errors/OmnixErrors.cs'
Require $errors 'constructor never writes TechnicalDetails to disk' 'automatic exception logging must not persist arbitrary technical details.'
Require $errors 'SafeLogText(friendlyMessage, 240)' 'friendly log text needs control-character and length sanitization.'
Forbid $errors '" | " + TechnicalDetails' 'TechnicalDetails must not be automatically written to gateway logs.'

Parse-Ps 'tools/provider-error-redaction-acceptance.ps1'
Require 'tools/provider-error-redaction-acceptance.ps1' 'PROVIDER-ERROR-REDACTION-RUNTIME-001' 'compiled redaction behavior needs a stable runtime TestId.'
Require 'tools/provider-error-redaction-acceptance.ps1' 'GatewayLogRedactedPass' 'runtime evidence must prove fake provider-body markers do not reach gateway logs.'
Require 'tools/provider-error-redaction-acceptance.ps1' 'PrivacyClassificationPass' 'OpenRouter policy detection must survive response-body redaction.'
Require 'tools/provider-error-redaction-acceptance.ps1' 'Guid.NewGuid()' 'redaction marker must be unpredictable and not a static fixture.'
Require '.github/workflows/request-budget.yml' 'provider-error-redaction-acceptance.ps1' 'runtime CI must execute compiled diagnostic redaction acceptance.'
Require '.github/workflows/request-budget.yml' 'PROVIDER-ERROR-REDACTION-RUNTIME-001' 'runtime CI must validate the redaction report TestId.'

$journal = 'src/OMNIX.Core/Logging/RuntimeDiagnosticJournal.cs'
Require $journal 'runtime-journey' 'human-readable runtime journey log must remain available.'
Require $journal 'runtime-events' 'structured JSONL runtime event log must remain available.'
Require $journal 'traceId' 'runtime events must carry request correlation identifiers.'
Require $journal 'SafeDetail' 'runtime detail metadata must pass through a redaction/sanitization boundary.'
Forbid $journal 'ArgumentsJson' 'tool arguments/document payload must never be written to runtime diagnostics.'
Forbid $journal 'UserTurn.Text' 'prompt text must never be written to runtime diagnostics.'
Require 'tools/runtime-diagnostic-journal-acceptance.ps1' 'RUNTIME-DIAGNOSTIC-JOURNAL-001' 'compiled journal behavior needs a stable runtime TestId.'
Require '.github/workflows/request-budget.yml' 'runtime-diagnostic-journal-acceptance.ps1' 'runtime CI must execute diagnostic journal acceptance.'
Require '.github/workflows/request-budget.yml' 'RUNTIME-DIAGNOSTIC-JOURNAL-001' 'runtime CI must validate the journal report TestId.'

if ($failures.Count -gt 0) {
    Write-Host 'OMNIX DIAGNOSTIC-REDACTION CONTRACT: FAIL' -ForegroundColor Red
    foreach ($f in $failures) { Write-Host " - $f" -ForegroundColor Red }
    exit 1
}
Write-Host 'OMNIX DIAGNOSTIC-REDACTION CONTRACT: PASS'
Write-Host 'Provider response bodies remain detection-only and are not emitted into mapped diagnostics or automatic logs.'
exit 0
