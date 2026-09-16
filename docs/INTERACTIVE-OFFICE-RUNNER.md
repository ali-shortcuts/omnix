# OMNIX Interactive Office Acceptance Runner

This document explains how to use the repository's manual `real-office-interactive` GitHub Actions workflow on a real Windows desktop with Excel, Word and PowerPoint installed.

This workflow is an **orchestration helper**, not a production approval shortcut. A PASS here does not replace restart, offline local-AI, live-provider, lifecycle, consumer-security, privacy UI, trusted signing, or the final production gate described in `PRODUCTION-EVIDENCE-RUNBOOK.md`.

## Why the runner must be interactive

Real Ribbon, WPF task-pane and UI Automation evidence is meaningful only in a normal signed-in Windows desktop session.

Do **not** run the OMNIX real Office acceptance runner as:

- a Windows service;
- LocalSystem;
- Session 0;
- a headless/background-only session with no Explorer shell;
- a machine where Excel/Word/PowerPoint are missing.

The read-only `tools/real-machine-preflight.ps1` fails closed when it detects those conditions.

## Runner requirements

Use a dedicated Windows x64 test machine or VM with:

- a normal interactive Windows user session;
- Windows Explorer running in that same session;
- desktop Excel, Word and PowerPoint installed;
- Git available for the Actions checkout;
- Windows PowerShell 5.1 available;
- enough rights for the normal OMNIX installer path, but no requirement to weaken Office or Windows security;
- the exact candidate installer already present at a known local path.

Configure a repository self-hosted GitHub Actions runner using GitHub's normal runner setup flow. Add the custom label:

```text
omnix-office-interactive
```

The workflow also requires the standard labels:

```text
self-hosted
Windows
X64
```

The runner process must be started **interactively under the same Windows user that will run Office**. Do not install/run it as a background Windows service for this acceptance workflow.

Self-hosted runners execute repository workflow code on the machine. Use a dedicated test environment, restrict repository write access, and do not keep unrelated secrets or personal documents on that runner.

## Before running the workflow

1. Sign in to Windows normally.
2. Confirm Explorer is running.
3. Close Excel, Word and PowerPoint.
4. Put the exact installer under test on the machine, for example:

```text
C:\OMNIX-Acceptance\OMNIX-Setup.exe
```

5. Compute its SHA-256:

```powershell
$installer = 'C:\OMNIX-Acceptance\OMNIX-Setup.exe'
(Get-FileHash -Algorithm SHA256 -LiteralPath $installer).Hash.ToLowerInvariant()
```

6. Make sure the source ref you intend to test corresponds to the installer's embedded `OMNIX-build-identity.json`. Do not mix a newer source checkout with an older installer.

## First run: preflight only

From GitHub Actions, manually dispatch:

```text
workflow: real-office-interactive
mode: preflight
installer_path: C:\OMNIX-Acceptance\OMNIX-Setup.exe
expected_installer_sha256: <64-hex SHA-256>
```

The preflight is read-only. It does not:

- launch Office;
- install or uninstall OMNIX;
- write Office registry state;
- restart Windows;
- change network/firewall state;
- change Defender or SmartScreen;
- change Office Trust Center;
- read Office documents.

It validates at least:

- Windows is the current OS;
- the runner is not Session 0;
- the current identity is not LocalSystem;
- Explorer is running in the same session;
- Excel, Word and PowerPoint are installed;
- Office processes are not already running;
- the installer exists and matches the expected SHA-256.

The result is written as `REAL-MACHINE-PREFLIGHT-001`.

Do not continue to Full Office E2E if preflight fails. Fix the environment instead of bypassing the check.

## Second run: canonical bound Full Office E2E

Only after preflight passes, manually dispatch the same workflow with:

```text
mode: full-office-e2e
```

Use the **same exact installer path and SHA-256**.

The workflow then calls the canonical production evidence wrapper:

```powershell
.\tools\bound-real-acceptance.ps1 `
  -Kind FullOfficeE2E `
  -InstallerPath <exact-installer> `
  -ExpectedInstallerSha256 <exact-sha256> `
  -SourceCommit <exact-checked-out-source>
```

After the Full Office E2E harness completes, the workflow runs `real-machine-preflight.ps1` again with `-RequireInstalledPayload`. That post-install check validates the installed `OMNIX-build-identity.json` and re-hashes the four primary assemblies:

```text
OMNIX.Core.dll
OMNIX.Excel.dll
OMNIX.Word.dll
OMNIX.PowerPoint.dll
```

## Coherent evidence-set validation

Individual bound JSON files are not enough. Before the workflow can complete successfully, it runs:

```powershell
.\tools\validate-bound-office-evidence.ps1
```

The validator does not trust any report as the source of the expected payload hashes. It independently re-reads the installed `OMNIX-build-identity.json`, verifies its exact source commit, and re-hashes all four installed primary assemblies. The resulting on-disk `CoreSha256` and `PayloadIdentitySha256` become the expected values for the report set.

The validator then treats these four reports as **one evidence set**:

```text
full-office-e2e.json
real-office-acceptance.json
real-office-ui-acceptance.json
taskpane-lifecycle-real-acceptance.json
```

It fails closed unless all four reports:

- exist and contain valid JSON;
- have `OverallPass=true`;
- carry `EvidenceBinding` schema 2 or newer;
- have `BuildIdentityTestId=OMNIX-BUILD-IDENTITY-001`;
- have `PrimaryAssembliesValidated=true`;
- are fresh under the configured evidence-age policy;
- carry the exact checked-out `SourceCommit`;
- carry the exact `CoreSha256` re-hashed from the installed payload;
- carry the exact `PayloadIdentitySha256` re-hashed from the installed build-identity file.

The Full Office E2E report must additionally prove that its installer SHA-256 equals the exact `expected_installer_sha256` supplied to the workflow and that `HashMatchedExpected=true`.

A mismatch in even one subordinate report rejects the whole evidence set. A consistent edit to all four JSON reports also fails if it no longer matches the installed files on disk. The validator produces the sanitized aggregate report:

```text
BOUND-OFFICE-EVIDENCE-SET-001
```

This closes both cross-report mixing and consistent-report tampering before the evidence is uploaded from the interactive runner.

## Uploaded evidence

The workflow stages only explicit JSON reports before artifact upload. It does not upload an arbitrary `%LOCALAPPDATA%` tree.

Expected staged reports include:

```text
real-machine-preflight-before.json
real-machine-preflight-after.json
bound-office-evidence-validation.json
full-office-e2e.json
real-office-acceptance.json
real-office-ui-acceptance.json
taskpane-lifecycle-real-acceptance.json
```

Artifacts are retained for the workflow's configured limited retention period. Acceptance reports are designed to contain status/hash/timing evidence rather than API keys or full private Office document content.

## What this workflow intentionally does not automate

The interactive workflow does **not** automate or manufacture these production conditions:

- Windows restart persistence;
- public-Internet-disconnected local-AI testing;
- live provider/account testing;
- repair/uninstall lifecycle phases;
- Defender + SmartScreen consumer-machine evidence;
- production Authenticode signing;
- final production approval.

Those remain explicit operator-controlled phases in `PRODUCTION-EVIDENCE-RUNBOOK.md`.

In particular, the workflow never restarts Windows and never disables or changes networking, Defender, SmartScreen, firewall, Office Trust Center or organizational policy to obtain a PASS.

## Production boundary

A successful interactive Full Office E2E run proves only the evidence classes actually exercised by that run and bound to that exact source/installed payload.

OMNIX is production-approved only when the complete evidence set for one exact final signed installer reaches:

```json
{
  "TestId": "OMNIX-FINAL-PRODUCTION-GATE-002",
  "FailureCount": 0,
  "OverallPass": true
}
```

Anything else remains non-production.

## Installer interaction during acceptance

Full Office E2E now opens the installer interactively by default. Keep the VM desktop unlocked and review the normal installer and Microsoft VSTO deployment prompts. Rejection or deployment failure must fail acceptance; do not import a root certificate or weaken Trust Center to obtain a PASS. The direct `tools/full-office-e2e.ps1 -SilentInstall` option is only for a candidate already trusted by that test profile. The canonical bound workflow retains interactive installation. A required restart remains a separate operator-controlled phase.
