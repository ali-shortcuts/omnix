# OMNIX conversation continuity anti-drift contract.
# Structural checks complement REQUEST-BUDGET-RUNTIME-001, which executes the compiled implementation.
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

function Require([string]$relative, [string]$needle, [string]$reason) {
    $text = Read-Repo $relative
    if ($text.IndexOf($needle, [StringComparison]::Ordinal) -lt 0) {
        $failures.Add("${relative}: missing '$needle' — $reason")
    }
}

function Forbid([string]$relative, [string]$needle, [string]$reason) {
    $text = Read-Repo $relative
    if ($text.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $failures.Add("${relative}: forbidden '$needle' — $reason")
    }
}

$modelPath = 'src/OMNIX.Core/AiGateway/Models/Models.cs'
$acceptancePath = 'tools/request-budget-acceptance.ps1'
$workflowPath = '.github/workflows/request-budget.yml'

# Preserve the existing hard provider-boundary limits.
Require $modelPath 'HardMaxHistoryTurns = 80' 'conversation continuity must not expand the hard turn cap.'
Require $modelPath 'HardMaxHistoryChars = 48 * 1024' 'conversation continuity must not expand the hard text cap.'
Require $modelPath 'MaxSingleHistoryTurnChars = 12 * 1024' 'one historical turn must remain bounded.'
Require $modelPath 'Images = null' 'historical image bytes must remain excluded from replay.'
Require $modelPath 'HistoricalImageMarker' 'omitted historical images must remain explicit to the model.'

# Continuity behavior: recent tail + first user anchor + relevant older turns, then chronological replay.
Require $modelPath 'RecentHistoryTurns = 20' 'a bounded recent tail must be retained.'
Require $modelPath 'continuityCharReserve' 'continuity anchors need a bounded reserve instead of consuming recent context.'
Require $modelPath 'maxChars / 4' 'continuity reserve must stay a bounded fraction of history budget.'
Require $modelPath 'FindFirstUserIndex' 'the first meaningful user goal must be eligible as a stable anchor.'
Require $modelPath 'ScoreContinuityRelevance' 'older turns must be selected by deterministic local relevance.'
Require $modelPath 'ExtractContinuityTerms' 'relevance selection must be local and deterministic.'
Require $modelPath 'ContainsInternalToolProtocol' 'stale internal tool protocol must not be promoted as conversation memory.'
Require $modelPath '.OrderBy(x => x.Key)' 'selected turns must be replayed in original chronological order.'
Require $modelPath 'source.UserTurn != null ? source.UserTurn.Text : null' 'the current request must drive older-turn relevance.'

# This layer is selection only: it must not create a hidden provider call or persistent memory store.
Forbid $modelPath 'SendAsync(' 'continuity selection must not make an additional AI/provider request.'
Forbid $modelPath 'HttpClient' 'continuity selection must remain network-free.'
Forbid $modelPath 'File.Write' 'continuity selection must not persist derived memory.'
Forbid $modelPath 'ProtectedData.Protect' 'continuity selection must not create a separate secret/state store.'

# Compiled runtime evidence must prove the three continuity guarantees and all prior safety caps.
foreach ($needle in @(
    'REQUEST-BUDGET-RUNTIME-001',
    'InitialUserAnchorPreservedPass',
    'RelevantOlderTurnPreservedPass',
    'RecentTailPreservedPass',
    'ContinuityOrderPass',
    'HistoricalImagesRemovedPass',
    'SourceRequestNotMutatedPass'
)) {
    Require $acceptancePath $needle 'conversation continuity behavior must remain covered by compiled runtime acceptance.'
}
Require $workflowPath 'Execute provider request budget acceptance' 'request-budget runtime acceptance must continue to run in CI.'
Require $workflowPath 'request-budget-acceptance.json' 'continuity runtime evidence must remain validated/uploaded through the request-budget report.'

if ($failures.Count -gt 0) {
    Write-Host 'OMNIX CONVERSATION-CONTINUITY CONTRACT: FAIL' -ForegroundColor Red
    foreach ($failure in $failures) { Write-Host " - $failure" -ForegroundColor Red }
    exit 1
}

Write-Host 'OMNIX CONVERSATION-CONTINUITY-CONTRACT-001: PASS'
Write-Host 'Recent-tail coherence, first-user anchoring, relevant older-turn recall, chronological replay and existing hard privacy/memory bounds are structurally intact.'
exit 0
