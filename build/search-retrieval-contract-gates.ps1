# OMNIX Office search retrieval anti-drift contract.
# Structural checks complement OFFICE-SEARCH-RETRIEVAL-RUNTIME-001.
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
    if ($text.IndexOf($needle,[StringComparison]::Ordinal) -lt 0) {
        $failures.Add("${relative}: missing '$needle' — $reason")
    }
}
function Forbid([string]$relative,[string]$needle,[string]$reason) {
    $text = Read-Repo $relative
    if ($text.IndexOf($needle,[StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $failures.Add("${relative}: forbidden '$needle' — $reason")
    }
}

$contracts = 'src/OMNIX.Core/Context/IHostAdapter.cs'
$tools = 'src/OMNIX.Core/Tools/Tools.cs'
$executor = 'src/OMNIX.Core/Tools/ToolExecutor.cs'
$search = 'src/OMNIX.Core/Context/SearchableOfficeAdapters.cs'
$coreProject = 'src/OMNIX.Core/OMNIX.Core.csproj'
$excelHost = 'src/OMNIX.Excel/ThisAddIn.cs'
$wordHost = 'src/OMNIX.Word/ThisAddIn.cs'
$pptHost = 'src/OMNIX.PowerPoint/ThisAddIn.cs'
$acceptance = 'tools/search-document-acceptance.ps1'
$workflow = '.github/workflows/search-retrieval.yml'
$gateway = 'src/OMNIX.Core/AiGateway/AiGateway.cs'

# Tool stays read-only and crosses the same cancellation/document-scope boundary as all reads.
Require $contracts 'IDocumentSearchProvider' 'targeted search must be an explicit optional read capability.'
Require $tools 'public const string SearchDocument = "search_document"' 'search_document must have one canonical tool name.'
Require $tools 'ReadSelection, ReadDocument, ReadPresentation, SearchDocument' 'search_document must be included with read tools.'
Require $executor 'case ToolNames.SearchDocument:' 'ToolExecutor must own search validation and scope enforcement.'
Require $executor 'adapter as IDocumentSearchProvider' 'search must require a search-capable Office adapter.'
Require $executor 'SearchQueryMaxChars = 200' 'search query size must be bounded before Office COM.'
Require $executor 'SearchResultMaxCount = 20' 'search result count must be bounded.'
Require $executor 'SearchResultCharCap = 6000' 'model-visible search output must be bounded.'
Require $executor 'EnsureRequestScope(ct);' 'search must share request cancellation/document-scope checks.'
Require $executor 'UntrustedData.Wrap("SEARCH_DOCUMENT RESULT", text)' 'search output must remain untrusted Office data.'

# Native bounded retrieval by host. Avoid replacing targeted search with whole-document materialization.
Require $search 'SearchableExcelHostAdapter' 'Excel needs a dedicated search-capable decorator.'
Require $search 'used.Find(' 'Excel retrieval must use native Range.Find.'
Require $search 'Excel.XlFindLookIn.xlValues' 'Excel must search values.'
Require $search 'Excel.XlFindLookIn.xlFormulas' 'Excel must search formulas as well as displayed values.'
Require $search 'EscapeExcelFindText' 'Excel wildcard characters must be escaped for literal targeted search.'
Require $search 'SearchableWordHostAdapter' 'Word needs a dedicated search-capable decorator.'
Require $search 'Word.WdFindWrap.wdFindStop' 'Word Find must stop rather than wrap forever.'
Require $search 'finder.MatchWildcards = false' 'Word targeted search must not interpret user text as wildcard syntax.'
Require $search 'doc.Range(nextStart, documentEnd)' 'Word search must advance instead of repeatedly matching one range.'
Require $search 'SearchablePowerPointHostAdapter' 'PowerPoint needs a dedicated search-capable decorator.'
Require $search 'range.Find(query' 'PowerPoint must use TextRange.Find for targeted text retrieval.'
Require $search 'shape.HasTable == Office.MsoTriState.msoTrue' 'PowerPoint search must include table cells.'
Require $search 'speaker notes' 'PowerPoint search must include speaker notes.'
Forbid $search 'doc.Content.Text' 'Word search must not materialize the whole document text.'
Forbid $search 'UsedRange.Value2' 'Excel search must not materialize the whole used range.'

# Thin Office hosts must actually instantiate the decorators.
Require $coreProject 'Context\SearchableOfficeAdapters.cs' 'search decorators must compile into OMNIX.Core.'
Require $excelHost 'new SearchableExcelHostAdapter' 'Excel must use the search-capable decorator.'
Require $wordHost 'new SearchableWordHostAdapter' 'Word must use the search-capable decorator.'
Require $pptHost 'new SearchablePowerPointHostAdapter' 'PowerPoint must use the search-capable decorator.'

# Models need to know targeted search exists; otherwise the runtime capability would rarely be used.
Require $gateway 'search_document' 'system prompt must advertise targeted document search.'
Require $gateway 'targeted' 'system prompt must distinguish targeted retrieval from broad read_document snapshots.'

# Runtime evidence must prove read-only classification, bounds, scope/cancellation and untrusted wrapping.
foreach ($needle in @(
    'OFFICE-SEARCH-RETRIEVAL-RUNTIME-001',
    'ReadOnlyPass',
    'ResultCountClampedPass',
    'ResultCharCapPass',
    'UntrustedWrapperPass',
    'FenceNeutralizedPass',
    'CancellationBeforeSearchPass',
    'ScopeLossBeforeSearchPass',
    'NoWritePathPass'
)) { Require $acceptance $needle 'compiled Office-search behavior must remain covered.' }
Require $workflow 'Execute Office search retrieval acceptance' 'compiled Office-search acceptance must run in CI.'
Require $workflow 'search-document-acceptance.json' 'sanitized Office-search evidence must be validated and uploaded.'

# No broad OS capability is allowed to enter through retrieval.
Forbid $search 'Process.Start' 'document search may not execute processes.'
Forbid $search 'Registry' 'document search may not access registry state.'
Forbid $search 'File.ReadAll' 'document search may not scan arbitrary files.'
Forbid $search 'HttpClient' 'document search must remain local to the active Office host.'

if ($failures.Count -gt 0) {
    Write-Host 'OMNIX OFFICE-SEARCH-RETRIEVAL CONTRACT: FAIL' -ForegroundColor Red
    foreach ($failure in $failures) { Write-Host " - $failure" -ForegroundColor Red }
    exit 1
}
Write-Host 'OMNIX OFFICE-SEARCH-RETRIEVAL-CONTRACT-001: PASS'
Write-Host 'Targeted search remains read-only, bounded, host-native, scope-aware, untrusted-data wrapped and free of broad OS/network capabilities.'
exit 0
