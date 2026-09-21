# OMNIX real Office functional acceptance
#
# Runs only on an interactive Windows machine with desktop Excel, Word and PowerPoint installed
# and OMNIX already installed. It creates temporary unsaved documents and proves that the compiled
# OMNIX.Core host adapters can read real Office context and that write tools remain fail-closed until
# confirmation. No user document is opened or saved. No Trust Center, Resiliency, registry policy,
# firewall, network adapter, or Office security setting is changed.

[CmdletBinding()]
param(
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\OMNIX",
    [string]$OutputPath = "$env:LOCALAPPDATA\OMNIX\logs\office-functional-acceptance.json",
    [int]$StartupDelayMs = 1200
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredFiles = @(
    'OMNIX.Core.dll',
    'Newtonsoft.Json.dll',
    'Microsoft.Office.Interop.Excel.dll',
    'Microsoft.Office.Interop.Word.dll',
    'Microsoft.Office.Interop.PowerPoint.dll'
)
foreach ($name in $requiredFiles) {
    $path = Join-Path $InstallDir $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required installed OMNIX file is missing: $path"
    }
}

$running = @()
foreach ($name in @('EXCEL','WINWORD','POWERPNT')) {
    if (Get-Process -Name $name -ErrorAction SilentlyContinue) { $running += $name }
}
if ($running.Count -gt 0) {
    throw "Close Excel, Word and PowerPoint before running this test. Currently running: $($running -join ', ')"
}

$logDir = Split-Path -Parent $OutputPath
if ($logDir) { New-Item -ItemType Directory -Force -Path $logDir | Out-Null }

function Load-OmnixAssembly([string]$name) {
    $path = Join-Path $InstallDir $name
    [void][Reflection.Assembly]::LoadFrom($path)
}

# Load dependencies explicitly from the installed payload so this test exercises the same binaries
# Office will use rather than a source-tree build output.
Load-OmnixAssembly 'Newtonsoft.Json.dll'
Load-OmnixAssembly 'Microsoft.Office.Interop.Excel.dll'
Load-OmnixAssembly 'Microsoft.Office.Interop.Word.dll'
Load-OmnixAssembly 'Microsoft.Office.Interop.PowerPoint.dll'
$corePath = Join-Path $InstallDir 'OMNIX.Core.dll'
[void][Reflection.Assembly]::LoadFrom($corePath)

# Small bridge used only to provide deterministic approve/deny delegates to ToolExecutor. It never
# touches Office or the network; it lets the real ToolExecutor prove the confirmation ordering.
$bridgeSource = @'
using System.Threading.Tasks;
using OMNIX.Core.Tools;
namespace OMNIX.E2E {
    public static class ConfirmationBridge {
        public static Task<bool> Approve(WritePreview preview) { return Task.FromResult(true); }
        public static Task<bool> Deny(WritePreview preview) { return Task.FromResult(false); }
    }
}
'@
Add-Type -TypeDefinition $bridgeSource -ReferencedAssemblies $corePath -ErrorAction Stop
$bridgeType = ('OMNIX.E2E.ConfirmationBridge' -as [type])
if ($null -eq $bridgeType) { throw 'Could not load OMNIX E2E confirmation bridge.' }

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } catch { }
    }
}

function New-Adapter([string]$typeName, [object[]]$args) {
    return New-Object -TypeName $typeName -ArgumentList $args
}

function New-ToolCall([string]$name, [string]$json) {
    $call = New-Object -TypeName 'OMNIX.Core.Tools.ToolCall'
    $call.Name = $name
    $call.ArgumentsJson = $json
    return $call
}

function Set-Confirmation($executor, [ValidateSet('None','Deny','Approve')] [string]$mode) {
    $prop = $executor.GetType().GetProperty('WriteConfirmation')
    if ($null -eq $prop) { throw 'ToolExecutor.WriteConfirmation property is missing.' }
    if ($mode -eq 'None') {
        $prop.SetValue($executor, $null, $null)
        return
    }
    $method = $bridgeType.GetMethod($mode)
    if ($null -eq $method) { throw "Confirmation bridge method not found: $mode" }
    $delegate = [Delegate]::CreateDelegate($prop.PropertyType, $method)
    $prop.SetValue($executor, $delegate, $null)
}

function Invoke-Tool($executor, $call, $adapter) {
    $task = $executor.ExecuteAsync($call, $adapter)
    return $task.GetAwaiter().GetResult()
}

function Contains-OrdinalIgnoreCase([string]$text, [string]$needle) {
    if ($null -eq $text -or $null -eq $needle) { return $false }
    return $text.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Wait-ProcessExit([string]$processName, [int]$timeoutMs = 5000) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (-not (Get-Process -Name $processName -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Milliseconds 200
    } while ($sw.ElapsedMilliseconds -lt $timeoutMs)
    return (-not (Get-Process -Name $processName -ErrorAction SilentlyContinue))
}

function Base-Result([string]$officeHostName) {
    return [ordered]@{
        Host = $officeHostName
        Installed = $true
        Started = $false
        Version = $null
        ContextReadPass = $false
        ReadToolPass = $false
        WriteNoConfirmationBlockedPass = $false
        WriteDeniedBlockedPass = $false
        WriteApprovedAppliedPass = $false
        VisionCaptureApplicable = $false
        VisionCapturePass = $null
        ProcessExitedCleanly = $false
        Error = $null
        Pass = $false
    }
}

function Test-Excel {
    $result = Base-Result 'Excel'
    $result.TypedCellWritePass = $false
    $result.ProfessionalFormatPass = $false
    $result.VisibleTargetPass = $false
    $app = $null; $book = $null; $sheet = $null; $selection = $null; $adapter = $null; $executor = $null
    $token = 'OMNIX_E2E_EXCEL_42'
    try {
        $app = New-Object -ComObject Excel.Application
        $result.Started = $true
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $result.Version = [string]$app.Version
        $book = $app.Workbooks.Add()
        $sheet = $app.ActiveSheet
        $sheet.Range('A1').Value2 = $token
        $sheet.Range('B1').Value2 = 20
        $sheet.Range('C1').Value2 = 22
        $sheet.Range('C2').Formula = '=B1+C1'
        $selection = $sheet.Range('A1:C2')
        [void]$selection.Select()
        Start-Sleep -Milliseconds $StartupDelayMs

        $adapter = New-Adapter 'OMNIX.Core.Context.ExcelHostAdapter' @($app, [Func[int]]{ 2000 }, [Func[int]]{ 6000 })
        $ctx = $adapter.ReadContext()
        $readSelection = [string]$adapter.ReadSelection()
        $result.ContextReadPass = [bool](
            $null -ne $ctx -and
            [string]$ctx.SelectionAddress -eq 'A1:C2' -and
            (Contains-OrdinalIgnoreCase ([string]$ctx.ValuesPreview) $token) -and
            (Contains-OrdinalIgnoreCase ([string]$ctx.FormulasPreview) '=B1+C1') -and
            (Contains-OrdinalIgnoreCase $readSelection $token))

        $executor = New-Object -TypeName 'OMNIX.Core.Tools.ToolExecutor'
        $readCall = New-ToolCall 'read_selection' '{}'
        $readResult = Invoke-Tool $executor $readCall $adapter
        $result.ReadToolPass = [bool]($readResult.Success -and (Contains-OrdinalIgnoreCase ([string]$readResult.ContentForModel) $token))

        $writeJson = '{"address":"A4","value":"OMNIX_WRITE_OK"}'
        $writeCall = New-ToolCall 'write_to_cell' $writeJson
        Set-Confirmation $executor 'None'
        $noneResult = Invoke-Tool $executor $writeCall $adapter
        $afterNone = [string]$sheet.Range('A4').Value2
        $result.WriteNoConfirmationBlockedPass = [bool]((-not $noneResult.Success) -and [string]::IsNullOrEmpty($afterNone))

        Set-Confirmation $executor 'Deny'
        $denyResult = Invoke-Tool $executor $writeCall $adapter
        $afterDeny = [string]$sheet.Range('A4').Value2
        $result.WriteDeniedBlockedPass = [bool]((-not $denyResult.Success) -and [string]::IsNullOrEmpty($afterDeny))

        Set-Confirmation $executor 'Approve'
        $approveResult = Invoke-Tool $executor $writeCall $adapter
        $afterApprove = [string]$sheet.Range('A4').Value2
        $result.WriteApprovedAppliedPass = [bool]($approveResult.Success -and $afterApprove -eq 'OMNIX_WRITE_OK')
        $typedCall = New-ToolCall 'write_to_cell' '{"address":"B4","value":42.5}'
        $typedResult = Invoke-Tool $executor $typedCall $adapter
        $typedValue = $sheet.Range('B4').Value2
        $result.TypedCellWritePass = [bool]($typedResult.Success -and [double]$typedValue -eq 42.5)

        $formatCall = New-ToolCall 'format_range' '{"address":"A4:B4","bold":true,"horizontalAlignment":"center","border":"thin","autofitColumns":true}'
        $formatResult = Invoke-Tool $executor $formatCall $adapter
        $formatted = $sheet.Range('A4:B4')
        $result.ProfessionalFormatPass = [bool]($formatResult.Success -and [bool]$formatted.Font.Bold)
        try { $result.VisibleTargetPass = [bool]([string]$app.Selection.Address($false,$false) -eq 'A4:B4') } catch { $result.VisibleTargetPass = $false }
    }
    catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147221164) {
            $result.Installed = $false
            $result.Error = 'Excel COM class is not registered.'
        } else { $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)" }
    }
    catch { $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)" }
    finally {
        if ($null -ne $book) { try { $book.Close($false) } catch { } }
        if ($null -ne $app) { try { $app.Quit() } catch { } }
        $adapter = $null; $executor = $null
        Release-ComObjectSafe $selection
        Release-ComObjectSafe $sheet
        Release-ComObjectSafe $book
        Release-ComObjectSafe $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $result.ProcessExitedCleanly = Wait-ProcessExit 'EXCEL'
    }
    $result.Pass = [bool]($result.Installed -and $result.Started -and $result.ContextReadPass -and $result.ReadToolPass -and
        $result.WriteNoConfirmationBlockedPass -and $result.WriteDeniedBlockedPass -and $result.WriteApprovedAppliedPass -and
        $result.TypedCellWritePass -and $result.ProfessionalFormatPass -and $result.VisibleTargetPass -and
        $result.ProcessExitedCleanly)
    return [pscustomobject]$result
}

function Test-Word {
    $result = Base-Result 'Word'
    $app = $null; $doc = $null; $range = $null; $adapter = $null; $executor = $null
    $token = 'OMNIX_E2E_WORD_ORIGINAL'
    try {
        $app = New-Object -ComObject Word.Application
        $result.Started = $true
        $app.Visible = $false
        $app.DisplayAlerts = 0
        $result.Version = [string]$app.Version
        $doc = $app.Documents.Add()
        $doc.Content.Text = "$token sample text"
        $range = $doc.Content
        [void]$range.Select()
        Start-Sleep -Milliseconds $StartupDelayMs

        $adapter = New-Adapter 'OMNIX.Core.Context.WordHostAdapter' @($app, [Func[int]]{ 6000 })
        $ctx = $adapter.ReadContext()
        $readSelection = [string]$adapter.ReadSelection()
        $result.ContextReadPass = [bool]($null -ne $ctx -and
            (Contains-OrdinalIgnoreCase ([string]$ctx.SelectionText) $token) -and
            (Contains-OrdinalIgnoreCase $readSelection $token))

        $executor = New-Object -TypeName 'OMNIX.Core.Tools.ToolExecutor'
        $readCall = New-ToolCall 'read_selection' '{}'
        $readResult = Invoke-Tool $executor $readCall $adapter
        $result.ReadToolPass = [bool]($readResult.Success -and (Contains-OrdinalIgnoreCase ([string]$readResult.ContentForModel) $token))

        $writeCall = New-ToolCall 'rewrite_selected_text' '{"text":"OMNIX_WORD_WRITE_OK"}'
        Set-Confirmation $executor 'None'
        [void](Invoke-Tool $executor $writeCall $adapter)
        $result.WriteNoConfirmationBlockedPass = Contains-OrdinalIgnoreCase ([string]$doc.Content.Text) $token

        [void]$doc.Content.Select()
        Set-Confirmation $executor 'Deny'
        [void](Invoke-Tool $executor $writeCall $adapter)
        $result.WriteDeniedBlockedPass = Contains-OrdinalIgnoreCase ([string]$doc.Content.Text) $token

        [void]$doc.Content.Select()
        Set-Confirmation $executor 'Approve'
        $approveResult = Invoke-Tool $executor $writeCall $adapter
        $result.WriteApprovedAppliedPass = [bool]($approveResult.Success -and
            (Contains-OrdinalIgnoreCase ([string]$doc.Content.Text) 'OMNIX_WORD_WRITE_OK'))
    }
    catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147221164) {
            $result.Installed = $false
            $result.Error = 'Word COM class is not registered.'
        } else { $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)" }
    }
    catch { $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)" }
    finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
        if ($null -ne $app) { try { $app.Quit() } catch { } }
        $adapter = $null; $executor = $null
        Release-ComObjectSafe $range
        Release-ComObjectSafe $doc
        Release-ComObjectSafe $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $result.ProcessExitedCleanly = Wait-ProcessExit 'WINWORD'
    }
    $result.Pass = [bool]($result.Installed -and $result.Started -and $result.ContextReadPass -and $result.ReadToolPass -and
        $result.WriteNoConfirmationBlockedPass -and $result.WriteDeniedBlockedPass -and $result.WriteApprovedAppliedPass -and
        $result.ProcessExitedCleanly)
    return [pscustomobject]$result
}

function Test-PowerPoint {
    $result = Base-Result 'PowerPoint'
    $result.VisionCaptureApplicable = $true
    $app = $null; $pres = $null; $slide = $null; $adapter = $null; $executor = $null
    $token = 'OMNIX_E2E_POWERPOINT_TITLE'
    try {
        $app = New-Object -ComObject PowerPoint.Application
        $result.Started = $true
        try { $app.Visible = -1 } catch { }
        $result.Version = [string]$app.Version
        $pres = $app.Presentations.Add()
        $slide = $pres.Slides.Add(1, 1) # ppLayoutTitle
        try { $slide.Shapes.Title.TextFrame.TextRange.Text = $token } catch { }
        try { $app.ActiveWindow.View.GotoSlide(1) } catch { }
        Start-Sleep -Milliseconds $StartupDelayMs

        $adapter = New-Adapter 'OMNIX.Core.Context.PowerPointHostAdapter' @($app, [Func[int]]{ 6000 })
        $ctx = $adapter.ReadContext()
        $readDocument = [string]$adapter.ReadDocument(6000)
        $result.ContextReadPass = [bool]($null -ne $ctx -and [int]$ctx.SlideCount -eq 1 -and
            (Contains-OrdinalIgnoreCase ([string]$ctx.SlideTitle) $token) -and
            (Contains-OrdinalIgnoreCase $readDocument $token))

        $executor = New-Object -TypeName 'OMNIX.Core.Tools.ToolExecutor'
        $readCall = New-ToolCall 'read_presentation' '{}'
        $readResult = Invoke-Tool $executor $readCall $adapter
        $result.ReadToolPass = [bool]($readResult.Success -and (Contains-OrdinalIgnoreCase ([string]$readResult.ContentForModel) $token))

        $visionCall = New-ToolCall 'capture_slide_as_image' '{"slide":"1"}'
        $visionResult = Invoke-Tool $executor $visionCall $adapter
        $result.VisionCapturePass = [bool]($visionResult.Success -and $null -ne $visionResult.CapturedPng -and $visionResult.CapturedPng.Length -gt 100)

        $insertCall = New-ToolCall 'insert_slide' '{"index":"2","title":"OMNIX_PPT_WRITE_OK","body":"temporary unsaved E2E slide"}'
        Set-Confirmation $executor 'None'
        [void](Invoke-Tool $executor $insertCall $adapter)
        $result.WriteNoConfirmationBlockedPass = ([int]$pres.Slides.Count -eq 1)

        Set-Confirmation $executor 'Deny'
        [void](Invoke-Tool $executor $insertCall $adapter)
        $result.WriteDeniedBlockedPass = ([int]$pres.Slides.Count -eq 1)

        Set-Confirmation $executor 'Approve'
        $approveResult = Invoke-Tool $executor $insertCall $adapter
        $title2 = ''
        if ([int]$pres.Slides.Count -ge 2) {
            try { $title2 = [string]$pres.Slides.Item(2).Shapes.Title.TextFrame.TextRange.Text } catch { }
        }
        $result.WriteApprovedAppliedPass = [bool]($approveResult.Success -and [int]$pres.Slides.Count -eq 2 -and
            (Contains-OrdinalIgnoreCase $title2 'OMNIX_PPT_WRITE_OK'))
    }
    catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147221164) {
            $result.Installed = $false
            $result.Error = 'PowerPoint COM class is not registered.'
        } else { $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)" }
    }
    catch { $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)" }
    finally {
        if ($null -ne $pres) { try { $pres.Close() } catch { } }
        if ($null -ne $app) { try { $app.Quit() } catch { } }
        $adapter = $null; $executor = $null
        Release-ComObjectSafe $slide
        Release-ComObjectSafe $pres
        Release-ComObjectSafe $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $result.ProcessExitedCleanly = Wait-ProcessExit 'POWERPNT'
    }
    $result.Pass = [bool]($result.Installed -and $result.Started -and $result.ContextReadPass -and $result.ReadToolPass -and
        $result.WriteNoConfirmationBlockedPass -and $result.WriteDeniedBlockedPass -and $result.WriteApprovedAppliedPass -and
        $result.VisionCapturePass -and $result.ProcessExitedCleanly)
    return [pscustomobject]$result
}

$results = New-Object System.Collections.Generic.List[object]
$results.Add((Test-Excel))
Start-Sleep -Milliseconds 600
$results.Add((Test-Word))
Start-Sleep -Milliseconds 600
$results.Add((Test-PowerPoint))

$installed = @($results | Where-Object { $_.Installed })
$requiredHostCountPass = ($installed.Count -eq 3)
$allFunctionalPass = $requiredHostCountPass -and (@($installed | Where-Object { -not $_.Pass }).Count -eq 0)
$allWritesGuarded = $requiredHostCountPass -and (@($installed | Where-Object {
    -not $_.WriteNoConfirmationBlockedPass -or -not $_.WriteDeniedBlockedPass -or -not $_.WriteApprovedAppliedPass
}).Count -eq 0)

$report = [ordered]@{
    TestId = 'OFFICE-FUNCTIONAL-REAL-001'
    EvidenceSchema = 1
    TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
    Windows = [Environment]::OSVersion.VersionString
    InteractiveSession = [Environment]::UserInteractive
    InstalledPayload = $InstallDir
    RequiredHosts = @('Excel','Word','PowerPoint')
    RequiredHostCountPass = $requiredHostCountPass
    AllWritesGuarded = $allWritesGuarded
    PowerPointVisionCapturePass = [bool](@($results | Where-Object { $_.Host -eq 'PowerPoint' })[0].VisionCapturePass)
    OverallPass = [bool]$allFunctionalPass
    Results = $results
    Safety = 'Temporary unsaved Office documents only; no user document, Trust Center, Resiliency, registry policy, firewall, network adapter, or Office security setting modified.'
}

$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$report | ConvertTo-Json -Depth 8

if (-not $report.OverallPass) { exit 1 }
exit 0