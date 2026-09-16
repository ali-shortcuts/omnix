// Shared by the production installer and the compiled regression harness.
// 0=ready, 1=restart requested by Microsoft, 2=repair/diagnose, 3=installer failure,
// 4=one recovery restart for an ambiguous success-without-runtime result.
function RuntimeDecision(const ExitCode: Integer; const Verified, RecoveryAlreadyRequested: Boolean): Integer;
begin
  if ExitCode = 3010 then Result := 1
  else if ExitCode <> 0 then Result := 3
  else if Verified then Result := 0
  else if RecoveryAlreadyRequested then Result := 2
  else Result := 4;
end;
