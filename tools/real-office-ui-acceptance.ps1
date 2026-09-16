# OMNIX real Office UI acceptance gate
#
# Run this on an interactive Windows desktop with Microsoft Office installed AFTER OMNIX.
# It verifies the part CI cannot prove: the OMNIX Ribbon tab is actually visible and the
# Open Workspace command actually opens the OMNIX task pane in Excel, Word and PowerPoint.
#
# Safety/correctness:
# - Uses normal Office COM automation + Windows UI Automation only.
# - Creates temporary blank Office files in memory and closes them without saving.
# - Does not alter Trust Center, Office Resiliency, registry policy, or security settings.
# - Does not click unrelated UI controls.
# - Workspace proof uses deterministic UI Automation IDs exposed by the OMNIX WPF UI; it never
#   accepts the already-visible Ribbon tab itself as evidence that the task pane opened.

[CmdletBinding()]
param(
    [string]$OutputPath = "$env:LOCALAPPDATA\OMNIX\logs\real-office-ui-acceptance.json",
    [int]$StartupDelayMs = 2200,
    [int]$WorkspaceDelayMs = 1600
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$hosts = @(
    [pscustomobject]@{ Name='Excel';      ProgId='Excel.Application';      Process='EXCEL' },
    [pscustomobject]@{ Name='Word';       ProgId='Word.Application';       Process='WINWORD' },
    [pscustomobject]@{ Name='PowerPoint'; ProgId='PowerPoint.Application'; Process='POWERPNT' }
)

$logDir = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } catch { }
    }
}

function Add-TemporaryDocument($app, [string]$hostName) {
    switch ($hostName) {
        'Excel' { return $app.Workbooks.Add() }
        'Word' { return $app.Documents.Add() }
        'PowerPoint' { return $app.Presentations.Add() }
        default { return $null }
    }
}

function Close-TemporaryDocument($doc, [string]$hostName) {
    if ($null -eq $doc) { return }
    try {
        switch ($hostName) {
            'Excel'      { $doc.Close($false) }
            'Word'       { $doc.Close(0) }
            'PowerPoint' { $doc.Close() }
        }
    } catch { }
}

function Get-OfficeWindowHandle($app, [string]$hostName) {
    try {
        switch ($hostName) {
            'Excel' { return [IntPtr]([int64]$app.Hwnd) }
            'Word' {
                if ($null -ne $app.ActiveWindow) { return [IntPtr]([int64]$app.ActiveWindow.Hwnd) }
            }
            'PowerPoint' {
                if ($null -ne $app.ActiveWindow) { return [IntPtr]([int64]$app.ActiveWindow.HWND) }
                return [IntPtr]([int64]$app.HWND)
            }
        }
    } catch { }
    return [IntPtr]::Zero
}

function Find-UiElementByExactName($root, [string]$name) {
    if ($null -eq $root) { return $null }
    try {
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, $name)
        return $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
    } catch { return $null }
}

function Find-UiElementByAutomationId($root, [string]$automationId) {
    if ($null -eq $root) { return $null }
    try {
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $automationId)
        return $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
    } catch { return $null }
}

function Find-UiElementByNameFragment($root, [string[]]$fragments) {
    if ($null -eq $root) { return $null }
    try {
        $trueCondition = [System.Windows.Automation.Condition]::TrueCondition
        $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $trueCondition)
        foreach ($element in $all) {
            try {
                $name = [string]$element.Current.Name
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                foreach ($fragment in $fragments) {
                    if ($name.IndexOf($fragment, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $element }
                }
            } catch { }
        }
    } catch { }
    return $null
}

function Test-UiElementVisible($element) {
    if ($null -eq $element) { return $false }
    try {
        if ([bool]$element.Current.IsOffscreen) { return $false }
        $rect = $element.Current.BoundingRectangle
        return ($rect.Width -ge 20 -and $rect.Height -ge 10)
    } catch { return $false }
}

function Activate-UiElement($element) {
    if ($null -eq $element) { return $false }

    $pattern = $null
    try {
        if ($element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) {
            ([System.Windows.Automation.SelectionItemPattern]$pattern).Select()
            return $true
        }
    } catch { }

    $pattern = $null
    try {
        if ($element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
            ([System.Windows.Automation.InvokePattern]$pattern).Invoke()
            return $true
        }
    } catch { }

    return $false
}

function Test-HostUi($officeHost) {
    $app = $null
    $doc = $null
    $started = Get-Date
    $result = [ordered]@{
        Host = $officeHost.Name
        Installed = $true
        Started = $false
        Version = $null
        WindowHandleFound = $false
        RibbonTabFound = $false
        RibbonTabActivated = $false
        OpenWorkspaceButtonFound = $false
        OpenWorkspaceInvoked = $false
        AutomaticWorkspaceVisible = $false
        WorkspaceEvidenceFound = $false
        WorkspaceEvidenceAutomationId = $null
        WorkspaceEvidenceName = $null
        WorkspaceEvidenceVisible = $false
        Error = $null
        DurationMs = 0
        Pass = $false
    }

    try {
        $app = New-Object -ComObject $officeHost.ProgId
        $result.Started = $true
        try { $app.Visible = $true } catch { }
        try { $app.DisplayAlerts = $false } catch { }

        $doc = Add-TemporaryDocument $app $officeHost.Name
        Start-Sleep -Milliseconds $StartupDelayMs
        try { $result.Version = [string]$app.Version } catch { $result.Version = 'unknown' }

        $hwnd = Get-OfficeWindowHandle $app $officeHost.Name
        if ($hwnd -eq [IntPtr]::Zero) { throw 'Could not obtain the active Office window handle.' }
        $result.WindowHandleFound = $true

        $root = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
        if ($null -eq $root) { throw 'UI Automation could not attach to the Office window.' }

        # Check before invoking the ribbon: a click must not mask broken automatic startup.
        $automaticDeadline = [DateTime]::UtcNow.AddSeconds(15)
        do {
            $automaticPane = Find-UiElementByAutomationId $root 'OMNIX.ChatInput'
            if ($null -eq $automaticPane) { $automaticPane = Find-UiElementByAutomationId $root 'OMNIX.WorkspaceRoot' }
            $result.AutomaticWorkspaceVisible = Test-UiElementVisible $automaticPane
            if ($result.AutomaticWorkspaceVisible) { break }
            Start-Sleep -Milliseconds 250
        } while ([DateTime]::UtcNow -lt $automaticDeadline)

        $tab = Find-UiElementByExactName $root 'OMNIX'
        if ($null -eq $tab) { $tab = Find-UiElementByNameFragment $root @('OMNIX') }
        $result.RibbonTabFound = ($null -ne $tab)
        if ($tab) {
            $result.RibbonTabActivated = Activate-UiElement $tab
            Start-Sleep -Milliseconds 500
        }

        $openButton = Find-UiElementByExactName $root 'Open Workspace'
        if ($null -eq $openButton) { $openButton = Find-UiElementByNameFragment $root @('Open Workspace') }
        $result.OpenWorkspaceButtonFound = ($null -ne $openButton)
        if ($openButton) {
            $result.OpenWorkspaceInvoked = Activate-UiElement $openButton
            Start-Sleep -Milliseconds $WorkspaceDelayMs
        }

        $workspace = Find-UiElementByAutomationId $root 'OMNIX.ChatInput'
        if ($null -eq $workspace) { $workspace = Find-UiElementByAutomationId $root 'OMNIX.WorkspaceRoot' }
        if ($null -eq $workspace) { $workspace = Find-UiElementByExactName $root 'OMNIX chat input' }

        if ($workspace) {
            $result.WorkspaceEvidenceVisible = Test-UiElementVisible $workspace
            try { $result.WorkspaceEvidenceAutomationId = [string]$workspace.Current.AutomationId } catch { }
            try { $result.WorkspaceEvidenceName = [string]$workspace.Current.Name } catch { }
            $result.WorkspaceEvidenceFound = [bool]$result.WorkspaceEvidenceVisible
        }

        $result.Pass = [bool](
            $result.Started -and
            $result.AutomaticWorkspaceVisible -and
            $result.WindowHandleFound -and
            $result.RibbonTabFound -and
            $result.RibbonTabActivated -and
            $result.OpenWorkspaceButtonFound -and
            $result.OpenWorkspaceInvoked -and
            $result.WorkspaceEvidenceFound)
    }
    catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147221164) {
            $result.Installed = $false
            $result.Error = 'Office COM class is not registered; this host appears not installed.'
        } else {
            $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)"
        }
    }
    catch {
        $result.Error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)"
    }
    finally {
        Close-TemporaryDocument $doc $officeHost.Name
        if ($null -ne $app) { try { $app.Quit() } catch { } }
        Release-ComObjectSafe $doc
        Release-ComObjectSafe $app
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $result.DurationMs = [int]((Get-Date) - $started).TotalMilliseconds
    }

    return [pscustomobject]$result
}

$running = @()
foreach ($officeHost in $hosts) {
    if (Get-Process -Name $officeHost.Process -ErrorAction SilentlyContinue) { $running += $officeHost.Name }
}
if ($running.Count -gt 0) {
    throw "Close Office before running the OMNIX UI acceptance gate. Currently running: $($running -join ', ')"
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($officeHost in $hosts) {
    $results.Add((Test-HostUi $officeHost))
    Start-Sleep -Milliseconds 900
}

$installed = @($results | Where-Object { $_.Installed })
$overallPass = ($installed.Count -eq 3) -and (@($installed | Where-Object { -not $_.Pass }).Count -eq 0)

$report = [ordered]@{
    TestId = 'OFFICE-UI-REAL-001'
    TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
    Machine = $env:COMPUTERNAME
    Windows = [Environment]::OSVersion.VersionString
    InteractiveSession = [Environment]::UserInteractive
    OverallPass = $overallPass
    Results = $results
}

$report | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
$report | ConvertTo-Json -Depth 8

if (-not $overallPass) { exit 1 }
exit 0