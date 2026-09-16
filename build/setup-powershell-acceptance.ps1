$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$folder = Join-Path $root 'build\artifact\setup helper test'
New-Item -ItemType Directory -Force $folder | Out-Null
$source = Get-Content (Join-Path $root 'installer\installer.iss') -Raw
$start = $source.IndexOf('procedure LogHelperOutput(')
$end = $source.IndexOf('function B2S(', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Production launcher missing.' }
$helper = $source.Substring($start, $end - $start)
$probe = Join-Path $folder 'probe.ps1'
'param([int]$Code=0); Write-Output "PROBE_RAN"; exit $Code' | Set-Content $probe
$log = Join-Path $folder 'helper.log'
$before = @(Get-ExecutionPolicy -List | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }) -join ';'
$iss = @"
[Setup]
AppName=OMNIX Setup Helper Test
AppVersion=1
DefaultDirName={tmp}\OMNIX-Helper-Test
CreateAppDir=no
Uninstallable=no
PrivilegesRequired=lowest
OutputDir=$folder
OutputBaseFilename=helper-test
[Code]
procedure InstallLog(const S: String);
begin
  SaveStringToFile('$log', S + #13#10, True);
end;
$helper
function InitializeSetup(): Boolean;
var Code: Integer;
begin
  if not Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -NonInteractive -File "$probe"', '', SW_HIDE, ewWaitUntilTerminated, Code) then
    RaiseException('Baseline launch failed');
  if Code = 0 then RaiseException('Restricted baseline did not block the script');
  if not RunSetupPowerShell('-NoProfile -NonInteractive -File "$probe"', Code) then
    RaiseException('Production helper did not launch');
  if Code <> 0 then RaiseException('Production helper did not execute the local script');
  if not RunSetupPowerShell('-NoProfile -NonInteractive -File "$probe" -Code 42', Code) then
    RaiseException('Failure probe did not launch');
  if Code <> 42 then RaiseException('Helper concealed child failure');
  if not RunSetupPowerShell('-NoProfile -NonInteractive -File "$folder\missing.ps1"', Code) then
    RaiseException('Missing-file probe did not launch');
  if Code = 0 then RaiseException('Missing script falsely passed');
  Result := True;
end;
"@
$issPath = Join-Path $folder 'helper-test.iss'
$iss | Set-Content $issPath -Encoding UTF8
$iscc = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
& $iscc /Q $issPath
if ($LASTEXITCODE -ne 0) { throw 'Could not compile production helper regression.' }
$old = $env:PSExecutionPolicyPreference
try {
    $env:PSExecutionPolicyPreference = 'Restricted'
    $run = Start-Process (Join-Path $folder 'helper-test.exe') -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait -PassThru
    if ($run.ExitCode -ne 0) { throw "Helper regression failed: $($run.ExitCode)" }
} finally { $env:PSExecutionPolicyPreference = $old }
$after = @(Get-ExecutionPolicy -List | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }) -join ';'
if ($before -ne $after) { throw 'Execution policy changed outside the helper process.' }
$text = Get-Content $log -Raw
if ($text -notmatch 'PROBE_RAN' -or $text -notmatch 'HELPER_EXIT: 42' -or $text -notmatch 'missing.ps1') { throw 'Child output/failure diagnostics missing.' }
@{ TestId='SETUP-POWERSHELL-001'; OverallPass=$true; RestrictedBaselineBlocked=$true; LocalScriptExecuted=$true; ChildFailurePreserved=$true; MissingScriptRejected=$true; PolicyUnchanged=$true; RealOfficeTested=$false } | ConvertTo-Json | Set-Content (Join-Path $root 'build\artifact\setup-powershell-acceptance.json')
Write-Host 'PASS: production Inno launcher tested with inherited Restricted policy, spaces in paths, child failure and missing script.'
