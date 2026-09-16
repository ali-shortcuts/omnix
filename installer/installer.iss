; ============================================================================
; OMNIX — per-user native Office installer
;
; Microsoft-documented VSTO discovery path used by this installer:
;   HKCU\Software\Microsoft\Office\<Application>\Addins\OMNIX
;
; IMPORTANT:
;   * Do NOT put VSTO registration under Office\16.0\... or Office\15.0\...
;   * Detect actual Excel/Word/PowerPoint executables before registration.
;   * Register only OMNIX-owned keys.
;   * Never clear Office Resiliency/DisabledItems/CrashingAddinList.
;   * Never change Trust Center or bypass Office policy.
; ============================================================================

#define MyAppName "OMNIX"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0-dev"
#endif
#define MyAppPublisher "Mr Ali"

[Setup]
AppId={{D5F28A04-617E-4C29-8F5D-A014E8C6B537}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\OMNIX
DisableDirPage=yes
DefaultGroupName=OMNIX
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=OMNIX-Setup-{#MyAppVersion}
SetupIconFile=..\build\omni.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
CloseApplications=no
UninstallDisplayIcon={app}\OMNIX.Excel.dll
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "payload\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion; Excludes: "vstor_redist.exe"
#if FileExists(AddBackslash(SourcePath) + "payload\vstor_redist.exe")
Source: "payload\vstor_redist.exe"; DestDir: "{tmp}"; Flags: dontcopy
#endif

[Icons]
Name: "{group}\Rescan Office Integration"; Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -File ""{app}\office-registration-maintenance.ps1"" -InstallDir ""{app}"""; WorkingDir: "{app}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
const
  CanonicalRegAddinsFmt = 'Software\Microsoft\Office\%0:s\Addins\OMNIX';
  LegacyRegAddinsFmt    = 'Software\Microsoft\Office\%0:s\%1:s\Addins\OMNIX';
  TrustedPubStore       = 'TrustedPublisher';
  RootStore             = 'Root';
  MaintenanceTaskName   = 'OMNIX Office Registration Maintenance';

var
  HostList: TStringList;
  InstallVerificationFailed: Boolean;
  VstoRestartNeeded: Boolean;
  MaintenanceTaskInstalled: Boolean;

#include "runtime-policy.iss"

procedure InstallLog(const Line: String);
var
  LogDir, Full: String;
begin
  try
    LogDir := ExpandConstant('{localappdata}') + '\OMNIX\logs';
    ForceDirectories(LogDir);
    Full := LogDir + '\install-debug.log';
    SaveStringToFile(Full, GetDateTimeString('yyyy-mm-dd hh:nn:ss', '-', ':') + '  ' + Line + #13#10, True);
  except
  end;
end;

function B2S(B: Boolean): String;
begin
  if B then Result := 'yes' else Result := 'no';
end;

function HostExeName(const Host: String): String;
begin
  if CompareText(Host, 'Excel') = 0 then Result := 'EXCEL.EXE'
  else if CompareText(Host, 'Word') = 0 then Result := 'WINWORD.EXE'
  else if CompareText(Host, 'PowerPoint') = 0 then Result := 'POWERPNT.EXE'
  else Result := '';
end;

function HostsSummary(): String;
var I: Integer;
begin
  Result := '';
  if HostList = nil then exit;
  for I := 0 to HostList.Count - 1 do
  begin
    if I > 0 then Result := Result + ', ';
    Result := Result + HostList[I];
  end;
end;

function IsProcessRunning(const ImageName: String): Boolean;
var ResultCode: Integer;
begin
  Exec(ExpandConstant('{cmd}'),
       '/C tasklist /FI "IMAGENAME eq ' + ImageName + '" | find /I "' + ImageName + '" >nul',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := (ResultCode = 0);
end;

function RegistryAppPathExists(const Exe: String): Boolean;
var P: String;
begin
  Result := False;
  if IsWin64 then
    if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\' + Exe, '', P) then
      if FileExists(RemoveQuotes(P)) then begin Result := True; exit; end;
  if RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\' + Exe, '', P) then
    if FileExists(RemoveQuotes(P)) then begin Result := True; exit; end;
  if RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\App Paths\' + Exe, '', P) then
    if FileExists(RemoveQuotes(P)) then Result := True;
end;

function IsHostInstalled(const Host: String): Boolean;
var Exe: String;
begin
  Exe := HostExeName(Host);
  Result := False;
  if Exe = '' then exit;

  if RegistryAppPathExists(Exe) then
  begin
    Result := True;
    exit;
  end;

  Result := FileExists(ExpandConstant('{pf32}\Microsoft Office\root\Office16\' + Exe)) or
    FileExists(ExpandConstant('{pf32}\Microsoft Office\Office16\' + Exe)) or
    FileExists(ExpandConstant('{pf32}\Microsoft Office\Office15\' + Exe));
  if (not Result) and IsWin64 then
    Result := FileExists(ExpandConstant('{pf64}\Microsoft Office\root\Office16\' + Exe)) or
      FileExists(ExpandConstant('{pf64}\Microsoft Office\Office16\' + Exe)) or
      FileExists(ExpandConstant('{pf64}\Microsoft Office\Office15\' + Exe));
end;

procedure DetectHosts();
begin
  if HostList <> nil then HostList.Free;
  HostList := TStringList.Create;
  if IsHostInstalled('Excel') then HostList.Add('Excel');
  if IsHostInstalled('Word') then HostList.Add('Word');
  if IsHostInstalled('PowerPoint') then HostList.Add('PowerPoint');
  InstallLog('Detected supported Office hosts: ' + HostsSummary());
end;

function RuntimeInView(const View: Integer; const KeyName: String): Boolean;
var Ver: String; Major: Integer;
begin
  Result := False;
  if not RegQueryStringValue(View, 'SOFTWARE\Microsoft\VSTO Runtime Setup\' + KeyName, 'Version', Ver) then exit;
  if Pos('.', Ver) = 0 then exit;
  Major := StrToIntDef(Copy(Ver, 1, Pos('.', Ver)-1), 0);
  if Major < 10 then exit;
  Result := True;
end;

function VstoRuntimeInstalled(): Boolean;
begin
  { Both Microsoft registration variants are valid; do not mistake v4 for missing runtime. }
  Result := RuntimeInView(HKLM32, 'v4R') or RuntimeInView(HKLM32, 'v4');
  if IsWin64 then Result := Result or RuntimeInView(HKLM64, 'v4R') or RuntimeInView(HKLM64, 'v4');
  InstallLog('VSTO Runtime present=' + B2S(Result));
  if Result then RegDeleteValue(HKCU, 'Software\OMNIX\Setup', 'RuntimeRecoveryRequested');
end;

procedure PreserveOfficeResiliencyState();
begin
  InstallLog('Office Resiliency state preserved unchanged.');
end;

procedure RemoveHostRegistration(const Host: String);
var Key: String;
begin
  Key := Format(CanonicalRegAddinsFmt, [Host]);
  if RegKeyExists(HKCU, Key) then
  begin
    RegDeleteKeyIncludingSubkeys(HKCU, Key);
    InstallLog('Removed canonical OMNIX key: HKCU\' + Key);
  end;

  Key := Format(LegacyRegAddinsFmt, ['16.0', Host]);
  if RegKeyExists(HKCU, Key) then RegDeleteKeyIncludingSubkeys(HKCU, Key);
  Key := Format(LegacyRegAddinsFmt, ['15.0', Host]);
  if RegKeyExists(HKCU, Key) then RegDeleteKeyIncludingSubkeys(HKCU, Key);
end;

procedure RemoveOwnedRegistration(const Key, Host: String);
var Actual, Expected: String;
begin
  Expected := ExpandConstant('{app}') + '\OMNIX.' + Host + '.vsto';
  StringChange(Expected, '\', '/');
  Expected := 'file:///' + Expected + '|vstolocal';
  if RegQueryStringValue(HKCU, Key, 'Manifest', Actual) then
    if CompareText(Actual, Expected) = 0 then RegDeleteKeyIncludingSubkeys(HKCU, Key)
    else InstallLog('UNINSTALL: preserved registration owned by another installation: ' + Key);
end;

procedure RemoveAddinRegistry();
var I: Integer; Host: String;
begin
  for I := 0 to 2 do
  begin
    if I = 0 then Host := 'Excel' else if I = 1 then Host := 'Word' else Host := 'PowerPoint';
    RemoveOwnedRegistration(Format(CanonicalRegAddinsFmt, [Host]), Host);
    RemoveOwnedRegistration(Format(LegacyRegAddinsFmt, ['16.0', Host]), Host);
    RemoveOwnedRegistration(Format(LegacyRegAddinsFmt, ['15.0', Host]), Host);
  end;
end;

function ManifestUri(const Host: String): String;
var P: String;
begin
  P := ExpandConstant('{app}') + '\OMNIX.' + Host + '.vsto';
  StringChange(P, '\', '/');
  Result := 'file:///' + P + '|vstolocal';
end;

function RegisterHost(const Host: String): Boolean;
var Key, Manifest, ReadBack: String; LoadReadBack: Cardinal;
begin
  Result := False;
  if not IsHostInstalled(Host) then exit;
  if not FileExists(ExpandConstant('{app}') + '\OMNIX.' + Host + '.vsto') then
  begin
    InstallLog('REGISTRATION_ERROR [' + Host + ']: deployment manifest is missing.');
    exit;
  end;

  RemoveHostRegistration(Host);
  Key := Format(CanonicalRegAddinsFmt, [Host]);
  Manifest := ManifestUri(Host);

  RegWriteStringValue(HKCU, Key, 'Description', 'OMNIX AI Office');
  RegWriteStringValue(HKCU, Key, 'FriendlyName', 'OMNIX');
  RegWriteDWordValue(HKCU, Key, 'LoadBehavior', 3);
  RegWriteStringValue(HKCU, Key, 'Manifest', Manifest);

  if not RegQueryStringValue(HKCU, Key, 'Manifest', ReadBack) then exit;
  if not RegQueryDWordValue(HKCU, Key, 'LoadBehavior', LoadReadBack) then exit;

  Result := (CompareText(ReadBack, Manifest) = 0) and (LoadReadBack = 3);
  InstallLog('Canonical VSTO registration [' + Host + '] path=HKCU\' + Key + ' manifest=' + ReadBack + ' load=' + IntToStr(LoadReadBack) + ' pass=' + B2S(Result));
end;

procedure RemoveMaintenanceTask();
var ResultCode: Integer; Helper: String;
begin
  Helper := ExpandConstant('{app}') + '\install-maintenance-task.ps1';
  if FileExists(Helper) then
  begin
    Exec('powershell.exe', '-NoProfile -NonInteractive -File "' + Helper + '" -InstallDir "' + ExpandConstant('{app}') + '" -Remove', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode = 0 then exit;
  end;
  Exec(ExpandConstant('{sys}') + '\schtasks.exe', '/Delete /TN "' + MaintenanceTaskName + '" /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function InstallMaintenanceTask(): Boolean;
var ResultCode: Integer; Helper: String;
begin
  Result := False;
  Helper := ExpandConstant('{app}') + '\install-maintenance-task.ps1';
  if not FileExists(Helper) then exit;
  Exec('powershell.exe', '-NoProfile -NonInteractive -File "' + Helper + '" -InstallDir "' + ExpandConstant('{app}') + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := (ResultCode = 0);
  InstallLog('Maintenance task install exit=' + IntToStr(ResultCode));
end;

function RunRegistrationMaintenance(): Boolean;
var ResultCode: Integer; ScriptPath: String;
begin
  Result := False;
  ScriptPath := ExpandConstant('{app}') + '\office-registration-maintenance.ps1';
  if not FileExists(ScriptPath) then exit;
  Exec('powershell.exe', '-NoProfile -NonInteractive -File "' + ScriptPath + '" -InstallDir "' + ExpandConstant('{app}') + '" -Quiet', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := (ResultCode = 0);
  InstallLog('Immediate registration maintenance exit=' + IntToStr(ResultCode));
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  HostList := nil;
  InstallVerificationFailed := False;
  VstoRestartNeeded := False;
  MaintenanceTaskInstalled := False;

  ForceDirectories(ExpandConstant('{localappdata}') + '\OMNIX\logs');
  InstallLog('=== OMNIX setup initialized ({#MyAppVersion}) ===');

  DetectHosts();
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var VstoExe: String; ResultCode, Decision: Integer; RecoveryRequested: Cardinal;
begin
  Result := '';
  NeedsRestart := False;
  try
    DetectHosts();
    if HostList.Count = 0 then
    begin
      Result := 'No supported desktop Excel, Word or PowerPoint installation was detected.';
      exit;
    end;
    if IsProcessRunning('excel.exe') or IsProcessRunning('winword.exe') or IsProcessRunning('powerpnt.exe') then
    begin
      Result := 'Close Excel, Word and PowerPoint, then retry installation. Save your work first.';
      exit;
    end;
    if not IsDotNetInstalled(net48, 0) then
    begin
      Result := 'Microsoft .NET Framework 4.8 or later is required. Install it from Microsoft, then run OMNIX Setup again.';
      exit;
    end;
    if VstoRestartNeeded then
    begin
      NeedsRestart := True;
      Result := 'Restart Windows to finish installing Microsoft VSTO Runtime, then run OMNIX Setup again.';
      exit;
    end;
    if not VstoRuntimeInstalled() then
    begin
      { Complete prerequisites before touching an existing OMNIX installation. }
      try ExtractTemporaryFile('vstor_redist.exe'); except end;
      VstoExe := ExpandConstant('{tmp}') + '\vstor_redist.exe';
      if not FileExists(VstoExe) then
      begin
        Result := 'The bundled Microsoft VSTO Runtime is missing. Download a complete OMNIX installer.';
        exit;
      end;
      if not ShellExec('runas', VstoExe, '/q /norestart', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
      begin
        Result := 'Microsoft VSTO Runtime installation could not start or was cancelled. The existing OMNIX installation was preserved.';
        exit;
      end;
      InstallLog('VSTO Runtime prerequisite exit=' + IntToStr(ResultCode));
      RecoveryRequested := 0;
      RegQueryDWordValue(HKCU, 'Software\OMNIX\Setup', 'RuntimeRecoveryRequested', RecoveryRequested);
      Decision := RuntimeDecision(ResultCode, VstoRuntimeInstalled(), RecoveryRequested <> 0);
      if (Decision = 1) or (Decision = 4) then
      begin
        RegWriteDWordValue(HKCU, 'Software\OMNIX\Setup', 'RuntimeRecoveryRequested', 1);
        VstoRestartNeeded := True;
        NeedsRestart := True;
        if Decision = 1 then
          Result := 'Restart Windows to finish installing Microsoft VSTO Runtime, then run OMNIX Setup again.'
        else
          Result := 'Microsoft reported success, but OMNIX cannot verify VSTO Runtime. Restart Windows once and retry. If verification still fails, repair Microsoft VSTO Runtime and inspect install-debug.log.';
        exit;
      end;
      if Decision = 2 then
      begin
        Result := 'VSTO Runtime remains unverified after a recovery restart was requested. If you already restarted, repair Microsoft VSTO Runtime and inspect install-debug.log. Repeated restarts are not a confirmed fix. The existing installation was preserved.';
        exit;
      end;
      if Decision <> 0 then
      begin
        Result := 'Microsoft VSTO Runtime installation failed (exit ' + IntToStr(ResultCode) + '). The existing OMNIX installation was preserved.';
        exit;
      end;
    end;
    InstallLog('PrepareToInstall hosts=[' + HostsSummary() + ']');
  except
    Result := 'OMNIX could not detect Microsoft Office: ' + GetExceptionMessage;
    InstallLog('OFFICE_DETECTION_ERROR: ' + GetExceptionMessage);
  end;
end;

function GetCustomSetupExitCode(): Integer;
begin
  if InstallVerificationFailed then Result := 10 else Result := 0;
end;

function UpdateReadyMemo(const Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := 'Office hosts detected: ' + HostsSummary() + NewLine +
            'VSTO registration: HKCU\Software\Microsoft\Office\<Host>\Addins\OMNIX' + NewLine +
            'Install folder: ' + ExpandConstant('{localappdata}') + '\Programs\OMNIX' + NewLine + NewLine +
            'OMNIX modifies only its own add-in registration. Office Trust Center and Resiliency remain unchanged.' + NewLine;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  I, ResultCode: Integer;
  AllOk: Boolean;
  CertPath, CertClassifier: String;
begin
  if CurStep = ssInstall then
  begin
    InstallLog('=== OMNIX install begin ===');
    InstallVerificationFailed := True;
    RemoveMaintenanceTask();
    PreserveOfficeResiliencyState();
    { Inno Setup owns file replacement and rollback; never delete the install tree here. }
  end;

  if CurStep = ssPostInstall then
  begin
    AllOk := True;

    CertPath := ExpandConstant('{app}') + '\install-vsto-trust.ps1';
    CertClassifier := '-NoProfile -File "' + CertPath + '" -InstallDir "' + ExpandConstant('{app}') + '" -HostNames "';
    for I := 0 to HostList.Count - 1 do
    begin
      if I > 0 then CertClassifier := CertClassifier + ',';
      CertClassifier := CertClassifier + HostList[I];
    end;
    CertClassifier := CertClassifier + '"';
    if WizardSilent then CertClassifier := CertClassifier + ' -Silent';
    if not FileExists(CertPath) then AllOk := False
    else if not Exec('powershell.exe', CertClassifier, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then AllOk := False
    else if ResultCode <> 0 then AllOk := False;
    { No addstore operation: Microsoft VSTOInstaller owns normal deployment trust. }

    if AllOk then
      for I := 0 to HostList.Count - 1 do
        if not RegisterHost(HostList[I]) then AllOk := False;

    if AllOk then
      if not RunRegistrationMaintenance() then AllOk := False;

    if AllOk then MaintenanceTaskInstalled := InstallMaintenanceTask();
    if not MaintenanceTaskInstalled then
      InstallLog('MAINTENANCE_WARNING: optional current-user re-scan task was not installed.');

    if AllOk and (not VstoRestartNeeded) then
    begin
      if FileExists(ExpandConstant('{app}') + '\post-install-verify.ps1') then
      begin
        if not Exec('powershell.exe', '-NoProfile -File "' + ExpandConstant('{app}') + '\post-install-verify.ps1"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
          AllOk := False;
        InstallLog('post-install-verify exit=' + IntToStr(ResultCode));
        if ResultCode <> 0 then AllOk := False;
      end
      else
      begin
        InstallLog('VERIFICATION_ERROR: post-install-verify.ps1 is missing.');
        AllOk := False;
      end;
    end;

    InstallVerificationFailed := not AllOk;
    InstallLog('=== OMNIX install end verified=' + B2S(AllOk) + ' restart=' + B2S(VstoRestartNeeded) + ' ===');

    if not AllOk then
      SuppressibleMsgBox('OMNIX files were installed, but Office add-in verification FAILED.' #13#10#13#10 +
             'See ' + ExpandConstant('{localappdata}') + '\OMNIX\logs\install-debug.log', mbError, MB_OK, IDOK)
    else if VstoRestartNeeded then
      SuppressibleMsgBox('OMNIX was installed. Restart Windows before the first Office test because the Microsoft VSTO Runtime requested a restart.', mbInformation, MB_OK, IDOK)
    else
      SuppressibleMsgBox('OMNIX was installed and registered for: ' + HostsSummary() + #13#10#13#10 +
             'Open Excel, Word or PowerPoint. The OMNIX Ribbon tab should load automatically.', mbInformation, MB_OK, IDOK);
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := not (IsProcessRunning('excel.exe') or IsProcessRunning('winword.exe') or IsProcessRunning('powerpnt.exe'));
  if not Result then SuppressibleMsgBox('Close Excel, Word and PowerPoint before uninstalling OMNIX.', mbError, MB_OK, IDOK);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
  CertMarker, DevThumbprintText: String;
  DevThumbprintRaw: AnsiString;
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveMaintenanceTask();
    RemoveAddinRegistry();
    PreserveOfficeResiliencyState();

    CertMarker := ExpandConstant('{app}') + '\dev-cert-thumbprint.txt';
    DevThumbprintRaw := '';
    DevThumbprintText := '';
    if FileExists(CertMarker) and LoadStringFromFile(CertMarker, DevThumbprintRaw) then
    begin
      DevThumbprintText := Trim(DevThumbprintRaw);
      if DevThumbprintText <> '' then
      begin
        Exec(ExpandConstant('{cmd}'), '/C certutil -user -delstore ' + TrustedPubStore + ' "' + DevThumbprintText + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
        Exec(ExpandConstant('{cmd}'), '/C certutil -user -delstore ' + RootStore + ' "' + DevThumbprintText + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      end;
    end;

    InstallLog('=== OMNIX uninstall: canonical + legacy OMNIX registration removed; Office Resiliency preserved ===');
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    if MsgBox('Also remove OMNIX settings and chat history (' + ExpandConstant('{localappdata}') + '\OMNIX)?' + #13#10 +
              'Choose Yes only if you do NOT plan to reinstall.', mbConfirmation, MB_YESNO) = IDYES then
      DelTree(ExpandConstant('{localappdata}') + '\OMNIX', True, True, True);
  end;
end;
