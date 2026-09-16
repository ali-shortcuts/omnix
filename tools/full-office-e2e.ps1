# OMNIX full real Office end-to-end acceptance orchestrator
#
# Intended for an interactive Windows desktop that has Excel + Word + PowerPoint installed.
# It optionally installs a supplied OMNIX installer, then runs:
#   1) automatic Office registration-maintenance/task safety acceptance,
#   2) strict two-launch COM automatic-load/persistence acceptance,
#   3) real Ribbon + Open Workspace + visible task-pane UI Automation acceptance,
#   4) repeated real Office task-pane close/reopen lifecycle acceptance in ALL 3 hosts,
#   5) real compiled OMNIX.Core Office-context/read/write/PowerPoint-Vision functional acceptance,
#   6) approved-but-invalid write boundary rejection in ALL 3 Office hosts,
#   7) real Office-context -> AiGateway/provider -> streaming WPF UI marker round-trip in ALL 3 hosts.
#
# This script does NOT restart Windows, alter networking, clear Office Resiliency, change Trust Center,
# or touch user Office documents. Reboot persistence remains a separate explicit before/after gate.

[CmdletBinding()]
param(
    [string]$InstallerPath,
    [string]$ExpectedInstallerSha256,
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\OMNIX",
    [string]$OutputPath = "$env:LOCALAPPDATA\OMNIX\logs\full-office-e2e.json",
    [switch]$SkipInstall,
    [switch]$SilentInstall,
    [switch]$SkipAiRoundTrip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logDir = Split-Path -Parent $OutputPath
if ($logDir) { New-Item -ItemType Directory -Force -Path $logDir | Out-Null }

function Assert-OfficeClosed {
    $running = @()
    foreach ($name in @('EXCEL','WINWORD','POWERPNT')) {
        if (Get-Process -Name $name -ErrorAction SilentlyContinue) { $running += $name }
    }
    if ($running.Count -gt 0) {
        throw "Close Excel, Word and PowerPoint before running the full OMNIX E2E test. Running: $($running -join ', ')"
    }
}

function Invoke-AcceptanceScript([string]$scriptName, [string]$reportPath, [string[]]$extraArgs = @()) {
    $scriptPath = Join-Path $scriptDir $scriptName
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw "Acceptance script missing: $scriptPath" }

    $args = @('-NoProfile','-File',$scriptPath,'-OutputPath',$reportPath) + $extraArgs
    $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
        throw "$scriptName did not produce its report: $reportPath (exit=$($p.ExitCode))"
    }
    $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
    return [pscustomobject]@{ ExitCode=$p.ExitCode; Report=$report }
}

Assert-OfficeClosed

$installerEvidence = [ordered]@{
    Requested = (-not $SkipInstall)
    SilentInstall = [bool]$SilentInstall
    Path = $null
    FileName = $null
    SizeBytes = $null
    Sha256 = $null
    ExpectedSha256 = if ([string]::IsNullOrWhiteSpace($ExpectedInstallerSha256)) { $null } else { $ExpectedInstallerSha256.ToLowerInvariant() }
    HashMatchedExpected = $null
    ExitCode = $null
    Pass = $false
    LogPath = $null
}

if ($SkipInstall) {
    if (-not (Test-Path -LiteralPath (Join-Path $InstallDir 'OMNIX.Core.dll') -PathType Leaf)) {
        throw "SkipInstall was requested but OMNIX is not installed at $InstallDir"
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedInstallerSha256)) {
        throw 'ExpectedInstallerSha256 cannot be proven when -SkipInstall is used. Final release E2E must test the exact installer.'
    }
    $installerEvidence.Pass = $true
}
else {
    if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
        throw 'InstallerPath is required unless -SkipInstall is supplied.'
    }
    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) { throw "Installer not found: $InstallerPath" }

    $installer = Get-Item -LiteralPath $InstallerPath
    if ($installer.Length -lt 1MB) { throw "Installer is unexpectedly small: $($installer.Length) bytes" }
    $installLog = Join-Path $logDir 'full-office-e2e-installer.log'
    if (Test-Path $installLog) { Remove-Item -LiteralPath $installLog -Force }

    $installerEvidence.Path = $installer.FullName
    $installerEvidence.FileName = $installer.Name
    $installerEvidence.SizeBytes = [int64]$installer.Length
    $installerEvidence.Sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $installer.FullName).Hash.ToLowerInvariant()
    $installerEvidence.LogPath = $installLog

    if (-not [string]::IsNullOrWhiteSpace($ExpectedInstallerSha256)) {
        $expected = $ExpectedInstallerSha256.Trim().ToLowerInvariant()
        if ($expected -notmatch '^[0-9a-f]{64}$') { throw 'ExpectedInstallerSha256 must be exactly 64 hexadecimal characters.' }
        $installerEvidence.HashMatchedExpected = [bool]($installerEvidence.Sha256 -eq $expected)
        if (-not $installerEvidence.HashMatchedExpected) {
            throw "Installer SHA256 mismatch. Expected=$expected Actual=$($installerEvidence.Sha256)"
        }
    }

    # Allow the operator to review Microsoft's normal VSTO deployment trust prompt.
    # Silent installation requires a candidate already trusted by this test profile.
    $installArgs = @('/NORESTART',('/LOG="' + $installLog + '"'))
    if ($SilentInstall) { $installArgs = @('/VERYSILENT','/SUPPRESSMSGBOXES') + $installArgs }
    $p = Start-Process -FilePath $installer.FullName -ArgumentList $installArgs -Wait -PassThru
    $installerEvidence.ExitCode = $p.ExitCode
    $installerEvidence.Pass = [bool]($p.ExitCode -eq 0 -and (Test-Path -LiteralPath (Join-Path $InstallDir 'OMNIX.Core.dll') -PathType Leaf))

    if (-not $installerEvidence.Pass) {
        $tail = ''
        if (Test-Path $installLog) {
            try { $tail = ((Get-Content -LiteralPath $installLog -Tail 40) -join [Environment]::NewLine) } catch { }
        }
        throw "OMNIX installer E2E stage failed (exit=$($p.ExitCode)). Log: $installLog`n$tail"
    }
}

Assert-OfficeClosed

$maintenancePath = Join-Path $logDir 'office-maintenance-real-acceptance.json'
$persistencePath = Join-Path $logDir 'real-office-acceptance.json'
$uiPath = Join-Path $logDir 'real-office-ui-acceptance.json'
$taskPaneLifecyclePath = Join-Path $logDir 'taskpane-lifecycle-real-acceptance.json'
$functionalPath = Join-Path $logDir 'office-functional-acceptance.json'
$writeBoundaryPath = Join-Path $logDir 'office-write-boundary-acceptance.json'
$aiPath = Join-Path $logDir 'real-office-ai-e2e.json'

$maintenance = Invoke-AcceptanceScript 'office-maintenance-real-acceptance.ps1' $maintenancePath @('-InstallDir',$InstallDir,'-RequiredHostCount','3')
Assert-OfficeClosed
$persistence = Invoke-AcceptanceScript 'real-office-acceptance.ps1' $persistencePath
Assert-OfficeClosed
$ui = Invoke-AcceptanceScript 'real-office-ui-acceptance.ps1' $uiPath
Assert-OfficeClosed
$taskPaneLifecycle = Invoke-AcceptanceScript 'taskpane-lifecycle-real-acceptance.ps1' $taskPaneLifecyclePath
Assert-OfficeClosed
$functional = Invoke-AcceptanceScript 'office-functional-acceptance.ps1' $functionalPath @('-InstallDir',$InstallDir)
Assert-OfficeClosed
$writeBoundary = Invoke-AcceptanceScript 'office-write-boundary-acceptance.ps1' $writeBoundaryPath @('-InstallDir',$InstallDir)
Assert-OfficeClosed

$ai = $null
if (-not $SkipAiRoundTrip) {
    $ai = Invoke-AcceptanceScript 'real-office-ai-e2e.ps1' $aiPath
    Assert-OfficeClosed
}

$officeVersions = @()
foreach ($name in @('Excel','Word','PowerPoint')) {
    $pRows = @($persistence.Report.Results | Where-Object { $_.Host -eq $name -and $_.Installed })
    $fRow = @($functional.Report.Results | Where-Object { $_.Host -eq $name -and $_.Installed }) | Select-Object -First 1
    $version = $null
    if ($pRows.Count -gt 0) { $version = [string]$pRows[0].Version }
    elseif ($null -ne $fRow) { $version = [string]$fRow.Version }
    $officeVersions += [ordered]@{ Host=$name; Version=$version }
}

$failures = New-Object System.Collections.Generic.List[string]
if (-not [bool]$installerEvidence.Pass) { $failures.Add('Installer stage failed.') }
if ($maintenance.ExitCode -ne 0 -or -not [bool]$maintenance.Report.OverallPass) { $failures.Add('Automatic Office registration maintenance/task acceptance failed.') }
if ($persistence.ExitCode -ne 0 -or -not [bool]$persistence.Report.OverallPass) { $failures.Add('Strict Office automatic-load/persistence acceptance failed.') }
if ($ui.ExitCode -ne 0 -or -not [bool]$ui.Report.OverallPass) { $failures.Add('Real Office Ribbon/workspace UI acceptance failed.') }
if ($taskPaneLifecycle.ExitCode -ne 0 -or -not [bool]$taskPaneLifecycle.Report.OverallPass) { $failures.Add('Real Office task-pane close/reopen lifecycle acceptance failed.') }
if ($functional.ExitCode -ne 0 -or -not [bool]$functional.Report.OverallPass) { $failures.Add('Real Office functional context/read/write/Vision acceptance failed.') }
if ($writeBoundary.ExitCode -ne 0 -or -not [bool]$writeBoundary.Report.OverallPass) { $failures.Add('Real Office approved-but-invalid AI write-boundary acceptance failed.') }
if (-not $SkipAiRoundTrip) {
    if ($null -eq $ai -or $ai.ExitCode -ne 0 -or -not [bool]$ai.Report.OverallPass) {
        $failures.Add('Real Office context -> AI Gateway/provider -> rendered UI marker round-trip failed.')
    }
}

$report = [ordered]@{
    TestId = 'OFFICE-E2E-REAL-001'
    EvidenceSchema = 5
    TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
    Windows = [Environment]::OSVersion.VersionString
    InteractiveSession = [Environment]::UserInteractive
    InstallDir = $InstallDir
    Installer = $installerEvidence
    OfficeVersions = $officeVersions
    Maintenance = [ordered]@{
        TestId = [string]$maintenance.Report.TestId
        ExitCode = [int]$maintenance.ExitCode
        OverallPass = [bool]$maintenance.Report.OverallPass
        MaintenanceTaskPresent = [bool]$maintenance.Report.MaintenanceTaskPresent
        MaintenanceTaskLimited = [bool]$maintenance.Report.MaintenanceTaskLimited
        MaintenanceTaskCurrentUser = [bool]$maintenance.Report.MaintenanceTaskCurrentUser
        MaintenanceTaskLogonTrigger = [bool]$maintenance.Report.MaintenanceTaskLogonTrigger
        InstalledHostCount = [int]$maintenance.Report.InstalledHostCount
        RegisteredHostCount = [int]$maintenance.Report.RegisteredHostCount
        AllInstalledRegistrationsPass = [bool]$maintenance.Report.AllInstalledRegistrationsPass
        SharedOfficeResiliencyPreserved = [bool]$maintenance.Report.SharedOfficeResiliencyPreserved
        OfficeProcessesRemainedClosed = [bool]$maintenance.Report.OfficeProcessesRemainedClosed
    }
    Persistence = [ordered]@{
        TestId = [string]$persistence.Report.TestId
        ExitCode = [int]$persistence.ExitCode
        OverallPass = [bool]$persistence.Report.OverallPass
        RequiredHostCountPass = [bool]$persistence.Report.RequiredHostCountPass
        TwoRoundsPerHostPass = [bool]$persistence.Report.TwoRoundsPerHostPass
        AutomaticLoadEveryRoundPass = [bool]$persistence.Report.AutomaticLoadEveryRoundPass
    }
    Ui = [ordered]@{
        TestId = [string]$ui.Report.TestId
        ExitCode = [int]$ui.ExitCode
        OverallPass = [bool]$ui.Report.OverallPass
    }
    TaskPaneLifecycle = [ordered]@{
        TestId = [string]$taskPaneLifecycle.Report.TestId
        ExitCode = [int]$taskPaneLifecycle.ExitCode
        OverallPass = [bool]$taskPaneLifecycle.Report.OverallPass
        InstalledHostCount = [int]$taskPaneLifecycle.Report.InstalledHostCount
        TwoRoundsPerHostRequired = [bool]$taskPaneLifecycle.Report.TwoRoundsPerHostRequired
    }
    Functional = [ordered]@{
        TestId = [string]$functional.Report.TestId
        ExitCode = [int]$functional.ExitCode
        OverallPass = [bool]$functional.Report.OverallPass
        RequiredHostCountPass = [bool]$functional.Report.RequiredHostCountPass
        AllWritesGuarded = [bool]$functional.Report.AllWritesGuarded
        PowerPointVisionCapturePass = [bool]$functional.Report.PowerPointVisionCapturePass
    }
    WriteBoundary = [ordered]@{
        TestId = [string]$writeBoundary.Report.TestId
        ExitCode = [int]$writeBoundary.ExitCode
        OverallPass = [bool]$writeBoundary.Report.OverallPass
        RequiredHostCountPass = [bool]$writeBoundary.Report.RequiredHostCountPass
        AllBoundaryChecksPass = [bool]$writeBoundary.Report.AllBoundaryChecksPass
    }
    AiRoundTrip = if ($SkipAiRoundTrip) {
        [ordered]@{ Required=$false; TestId=$null; ExitCode=$null; OverallPass=$null; AllMarkerRoundTripsPass=$null }
    } else {
        [ordered]@{
            Required = $true
            TestId = [string]$ai.Report.TestId
            ExitCode = [int]$ai.ExitCode
            OverallPass = [bool]$ai.Report.OverallPass
            RequiredHostCountPass = [bool]$ai.Report.RequiredHostCountPass
            AllMarkerRoundTripsPass = [bool]$ai.Report.AllMarkerRoundTripsPass
            AllProcessesExitedPass = [bool]$ai.Report.AllProcessesExitedPass
        }
    }
    FailureCount = $failures.Count
    Failures = @($failures)
    OverallPass = ($failures.Count -eq 0)
    RemainingSeparateReleaseGates = @(
        'Windows restart persistence before/after an actual user-initiated restart',
        'offline local-model round-trip with public Internet disconnected',
        'full live provider matrix/rate-limit/privacy/Vision evidence',
        'repair/reinstall/uninstall lifecycle preservation evidence',
        'consumer-machine Defender/SmartScreen with normal protections enabled',
        'trusted production Authenticode signature'
    )
    Safety = 'Temporary unsaved Office files only. Automatic maintenance is per-user/limited and may change only OMNIX-owned Addins keys; no automatic restart/network/firewall/Trust Center/Office Resiliency manipulation.'
}

$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$report | ConvertTo-Json -Depth 10

if (-not $report.OverallPass) { exit 1 }
exit 0
