# Test Office integration without changing the primary computer

## What the research establishes

A VSTO desktop add-in needs Windows, the VSTO runtime and desktop Office. Linux unit tests, browser Office, an EXE build or a registry-only test cannot prove that Office renders a ribbon and WPF pane. A separate Windows virtual machine can test this without installing OMNIX on the primary computer.

Use a dedicated Windows 11 x64 VM with appropriately licensed desktop Office, an interactive signed-in user and normal security settings. Hyper-V supports checkpoints, so each clean-install or repair scenario can start from a known baseline. A persistent VM is preferable to a disposable sandbox for before/after Windows-restart evidence. Windows Sandbox is useful for a quick isolated check, but does not by itself supply Office or replace the persistent reboot matrix.

Microsoft documents that Office assumes an interactive desktop and that noninteractive service automation can fail or hang. Start the existing GitHub self-hosted runner using `run.cmd` in the signed-in desktop, not as a Windows service or LocalSystem. Use the existing `omnix-office-interactive` label and `real-office-interactive` workflow. The repository runner inventory checked during this work contained zero runners; no remote VM or licensed Office session is currently connected to this workspace.

## Concrete execution order

1. Prepare separate VM snapshots for Microsoft 365 x64 and x86 Office; test older Office only where a corresponding environment is available. Record exact Windows/Office versions and architecture. Do not claim untested combinations.
2. Install/licence Office inside the VM, finish first-run prompts, close all Office processes, and create a clean checkpoint.
3. Download the exact candidate installer, manifest, complete-source ZIP and provenance from the same release. Check hashes and check out that exact source commit.
4. Run `tools/real-machine-preflight.ps1` or dispatch the existing interactive workflow in `preflight` mode. It must reject Session 0, missing Office or wrong payload identity.
5. Run `tools/bound-real-acceptance.ps1 -Kind FullOfficeE2E` with the exact installer and source binding, or use the workflow's `full-office-e2e` mode. The updated UI test now requires a visible workspace BEFORE invoking Open Workspace.
6. Follow `PRODUCTION-EVIDENCE-RUNBOOK.md` for RestartBefore/RestartAfter, approved live-provider testing, local offline testing, repair/uninstall lifecycle and consumer security. Restart the VM, not the primary computer. Keep screenshots and bounded sanitized reports, not private documents or provider keys.
7. Restore the clean checkpoint before the next scenario. Keep evidence bound to the same source and installer hashes. A rebuild changes the candidate and invalidates a blanket approval for the previous binary.

## Required scenarios

- Clean machine without VSTO: prerequisite success, 3010, exit 0 with unverified runtime, repeated unverified runtime, and real nonzero failure.
- Existing VSTO installed through v4 and v4R registrations.
- Office open during installation: refuse without modifying the existing application.
- Word, Excel, PowerPoint: automatic add-in load and visible workspace without a click; repeated Open Workspace calls; close/reopen pane; multiple documents; cancel document close; close/reopen the application.
- Upgrade, repair, uninstall and Windows reboot; verify the exact installed assemblies and startup behavior afterward.
- Trust prompts, locked desktop, protected documents and security blocking must be reported as failures/limitations, never bypassed to obtain a PASS.

## Primary sources

- [Microsoft VSTO runtime overview](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime-overview?view=visualstudio)
- [Microsoft Office automation requirements](https://support.microsoft.com/en-us/visio/considerations-for-server-side-automation-of-office)
- [Microsoft Hyper-V checkpoints](https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/checkpoints)
- [Microsoft Windows Sandbox](https://learn.microsoft.com/en-us/windows/security/application-security/application-isolation/windows-sandbox/)
- [GitHub self-hosted runners](https://docs.github.com/actions/hosting-your-own-runners)
- [Microsoft runtime download, version 10.0.60910](https://www.microsoft.com/en-us/download/details.aspx?id=105522)
