$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=Split-Path -Parent $PSScriptRoot
$folder=Join-Path $root 'build\artifact\runtime-policy-test'
New-Item -ItemType Directory -Force $folder | Out-Null
$policy=Join-Path $root 'installer\runtime-policy.iss'
$iscc=Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
$source=@"
[Setup]
AppName=OMNIX Runtime Policy Test
AppVersion=1
DefaultDirName={tmp}\OMNIX-Runtime-Policy
CreateAppDir=no
Uninstallable=no
PrivilegesRequired=lowest
OutputDir=$folder
OutputBaseFilename=runtime-policy-test
[Code]
#include "$policy"
procedure Check(const Code: Integer; const Verified, Retried: Boolean; const Expected: Integer);
begin
  if RuntimeDecision(Code, Verified, Retried) <> Expected then
    RaiseException('Runtime prerequisite decision regression');
end;
function InitializeSetup(): Boolean;
begin
  Check(0, True, False, 0);
  Check(0, True, True, 0);
  Check(3010, True, False, 1);
  Check(3010, False, True, 1);
  Check(0, False, False, 4);
  Check(0, False, True, 2);
  Check(1603, False, False, 3);
  Check(1603, True, True, 3);
  Result := True;
end;
"@
$iss=Join-Path $folder 'test.iss';$source | Set-Content -LiteralPath $iss -Encoding UTF8
& $iscc /Q $iss
if($LASTEXITCODE -ne 0){throw 'Could not compile the actual installer runtime decision logic.'}
$run=Start-Process (Join-Path $folder 'runtime-policy-test.exe') -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait -PassThru
if($run.ExitCode -ne 0){throw "Runtime policy regression failed: $($run.ExitCode)"}
@{TestId='VSTO-RUNTIME-POLICY-001';Cases=8;OverallPass=$true;RealOfficeTested=$false} | ConvertTo-Json | Set-Content (Join-Path $root 'build\artifact\runtime-policy-acceptance.json')
Write-Host 'PASS: eight cases executed from the production Inno runtime decision function.'
