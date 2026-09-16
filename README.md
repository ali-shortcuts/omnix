> **Canonical development:** use `main`. The separate `rebuild/native-v4` preview line is superseded; see [reconciliation](docs/UNIFIED-DEVELOPMENT.md). For installation tests without changing your primary computer, see the [isolated Office VM lab](docs/ISOLATED-OFFICE-LAB.md). Real Office/reboot acceptance remains required.

# OMNIX — Native AI Bridge for Microsoft Office

OMNIX is a Windows Office AI bridge: a native **C# / WPF / VSTO** add-in that connects **Excel, Word and PowerPoint** to local AI runtimes, cloud providers and custom OpenAI-compatible endpoints from one docked workspace inside Office.

> **Release status:** active v3 rebuild. The solution and development installer are built and hardened in Windows CI, but this repository does **not** call the current branch production-ready until the real-machine release gates in [issue #53](https://github.com/ali-shortcuts/OMNIX-v2/issues/53) pass. Green hosted CI is not a substitute for real desktop Excel/Word/PowerPoint.

Download the newest **Development Preview** `.exe` from [GitHub Releases](https://github.com/ali-shortcuts/OMNIX-v2/releases). Development previews are intentionally marked `DEVELOPMENT_ONLY`; they are not production releases. This is the canonical native Office repository. The separate `OMINIX.exe` repository contains an older browser/server prototype; its executable is not this native installer. See the [repository audit and consolidation decision](docs/REPOSITORY-AUDIT-2026-09-10.md).

The installer checks Office, .NET Framework 4.8 and VSTO prerequisites before replacing existing files. Close all Office applications before installing. A prerequisite restart stops the upgrade so the existing installation is preserved. Post-install verification failure returns exit code `10`, including in silent mode.

In Settings, a manually entered model is kept even if a server has no model catalog. **Test Connection** sends a small text request to that selected model, subject to the configured privacy mode; provider usage charges may apply. A successful text test does not claim Vision support.

## Product contract

```text
Microsoft Office
  Excel / Word / PowerPoint
        │
        ▼
Native OMNIX Ribbon + docked WPF Workspace
        │
        ▼
Office Context Engine
  structured selection/document/slide context
  + bounded visual capture when needed
        │
        ▼
Privacy Gate + Context Limiter + Whitelisted Tool Executor
        │
        ▼
OMNIX AI Gateway
        │
        ├── Local AI: Ollama / LM Studio
        ├── Cloud: Gemini / Groq / OpenRouter / Mistral / Hugging Face / Cerebras
        └── Custom OpenAI-compatible endpoint
```

OMNIX is the bridge, not the destination. The user works in Office; OMNIX obtains the minimum relevant Office context, routes it according to the privacy policy, streams the model response back into the side panel and exposes only narrowly scoped Office actions.

## What the rebuilt version contains

- **Three native Office hosts** — `OMNIX.Excel`, `OMNIX.Word`, `OMNIX.PowerPoint` share one `OMNIX.Core`.
- **No browser/Web UI/Node.js dependency** for normal operation. The main UI is WPF hosted in a VSTO Custom Task Pane.
- **Compact docked workspace** — designed around a narrow Office side panel rather than a web page squeezed into Office.
- **Per-window isolation** — each Office document window owns its own workspace controller, AI gateway, provider state, cancellation token, cloud-consent session and conversation context.
- **Streaming + cancellation** — provider responses stream into the rendered chat; Stop cancels the active request.
- **Structured Office context** — bounded Excel ranges/values/formulas, Word selection/document structure, and PowerPoint slide/presentation text are read through host adapters.
- **Vision** — current Office view/selection capture, Excel chart capture and PowerPoint slide capture can be attached to Vision-capable models. The model is explicitly told not to claim visibility outside the captured/structured scope.
- **Local AI** — Ollama and LM Studio are discovered in the background and can be preferred when available.
- **Cloud/custom providers** — Gemini, Groq, OpenRouter, Mistral AI, Hugging Face Inference Providers, Cerebras and custom OpenAI-compatible endpoints.
- **Dynamic provider behavior** — live model discovery is used where supported. OpenRouter runtime acceptance prefers `openrouter/free` and current `:free` routes before paid routes. Free/free-tier/trial status is treated as provider/account dependent rather than guaranteed forever.
- **Privacy modes** — `Local Only`, `Cloud Allowed`, `Ask Before Sending`. Enforcement occurs in the AI Gateway before a cloud provider's send path, not just in the UI.
- **DPAPI-protected API keys** — provider keys are protected for the current Windows user in `%LOCALAPPDATA%\OMNIX\settings.dat`; plaintext keys are not written to logs or reports.
- **Encrypted bounded history** — chat history is protected for the current Windows user, bounded by configured age/count limits, and raw attached Office screenshots are not persisted into history.
- **Untrusted Office data boundary** — document content is data, never executable instructions. Prompt-like content from a workbook/document/presentation remains inside the untrusted-data boundary.
- **Whitelisted Office tools only** — no unrestricted PowerShell, CMD, registry, process-control or arbitrary filesystem tool is exposed to the model.
- **Fail-closed writes** — a write tool requires a before/after preview and explicit user approval. If the confirmation callback is unavailable or the user denies the preview, no write is applied.
- **Categorized errors** — network/auth/model/privacy/provider/local-runtime failures are kept distinct instead of turning every failure into “check your Internet.”

## Provider matrix

| Provider | Kind | API key | Notes |
| --- | --- | --- | --- |
| Ollama | Local | No | Local models discovered from the runtime |
| LM Studio | Local | No | OpenAI-compatible local runtime |
| Google Gemini | Cloud | Yes | Model/capability availability is account/API dependent |
| Groq | Cloud | Yes | Free-plan eligibility/limits are account dependent |
| OpenRouter | Cloud | Yes | Supports live discovery; free router/free routes are preferred when available |
| Mistral AI | Cloud | Yes | Account/free-mode eligibility can change |
| Hugging Face Inference Providers | Cloud | Yes | Routes/models/credits are account and provider dependent |
| Cerebras | Cloud | Yes | Access/trial/quota are account dependent |
| Custom OpenAI-compatible | Local or Cloud | Optional | User supplies base URL, model and optional key |

OMNIX does not hard-code a permanent promise that a third-party cloud model is “free.” Provider availability, quotas and pricing can change independently of this repository.

## Office integration scope

The native v3 architecture targets **Windows desktop Microsoft Office environments that support VSTO**. The installer detects the Office generation/platform and only registers hosts found on the machine. Compatibility must be reported from evidence, not assumed.

Use these support states when documenting a tested environment:

```text
UNSUPPORTED / LEGACY / PARTIAL / SUPPORTED / FULLY_TESTED
```

`FULLY_TESTED` is reserved for a specific Office/Windows environment with real-machine evidence. OMNIX does not claim that one VSTO build universally supports every historical Office version or non-Windows Office.

## Installation and build identity

The development installer is an Inno Setup per-user installer. Its supported path is:

```text
intentional user install
  → detect Office / host apps / platform
  → ensure official Microsoft VSTO Runtime prerequisite when genuinely required
  → copy validated Excel + Word + PowerPoint + Core payload
  → install OMNIX-build-identity.json
  → register only OMNIX-owned Office add-in keys
  → preserve shared Office Resiliency state
  → perform post-install diagnostics
```

`OMNIX-build-identity.json` is generated during packaging for the exact source commit and records SHA-256 for `OMNIX.Core.dll`, `OMNIX.Excel.dll`, `OMNIX.Word.dll` and `OMNIX.PowerPoint.dll`. Post-install verification and the real-machine evidence binder reject missing or mismatched identity/assembly hashes.

OMNIX does **not** clear shared `DisabledItems`/`CrashingAddinList` state to force itself enabled. It does not bypass Trust Center or organizational Office policy.

### Development manifest trust

Current CI development builds use a temporary development manifest certificate. The installer only performs development trust handling when the bundled public certificate is classified as self-signed, records its exact thumbprint, and uninstall targets only that recorded development thumbprint. This is **development-only**, not the production trust model.

Production release requires a normal trusted code-signing certificate and valid timestamped Authenticode evidence. `build/sign-production.ps1` signs using an already provisioned certificate in the Windows certificate/key provider; it does not create/export a private key or handle a PFX password.

## Canonical real release gates

The authoritative operator procedure is [docs/PRODUCTION-EVIDENCE-RUNBOOK.md](docs/PRODUCTION-EVIDENCE-RUNBOOK.md).

For Office/provider/offline/restart production evidence, the canonical entrypoint is **only**:

```text
tools/bound-real-acceptance.ps1 -Kind FullOfficeE2E
tools/bound-real-acceptance.ps1 -Kind Provider
tools/bound-real-acceptance.ps1 -Kind LocalOffline
tools/bound-real-acceptance.ps1 -Kind RestartBefore
tools/bound-real-acceptance.ps1 -Kind RestartAfter
```

The bound runner executes the underlying real harness, requires a freshly generated report, validates the installed `OMNIX-build-identity.json`, validates Core + Excel + Word + PowerPoint assembly hashes, and only then adds `EvidenceBinding` with the exact source commit and installed payload identity.

Lifecycle uses the same identity model: Baseline validates the installed source/payload and all four primary assemblies, AfterRepair must reproduce the same `SourceCommit`, `CoreSha256` and `PayloadIdentitySha256`, and AfterUninstall carries the verified binding into the final report after the application files are correctly removed.

The lower-level scripts below are **implementation harnesses**. Their raw PASS files are not production evidence by themselves:

- `tools/real-office-acceptance.ps1`
- `tools/real-office-ui-acceptance.ps1`
- `tools/office-functional-acceptance.ps1`
- `tools/real-office-ai-e2e.ps1`
- `tools/full-office-e2e.ps1`
- `tools/reboot-persistence-acceptance.ps1`
- `tools/local-offline-acceptance.ps1`
- `tools/provider-acceptance.ps1`

Other canonical release entrypoints are:

- `tools/privacy-acceptance.ps1` — deterministic compiled Gateway privacy enforcement.
- `tools/lifecycle-acceptance.ps1` — exact-installer + installed-payload-bound Baseline → AfterRepair → AfterUninstall lifecycle evidence.
- `tools/consumer-security-acceptance.ps1` — real consumer Defender + SmartScreen evidence without disabling/bypassing protection.
- `tools/release-readiness.ps1` — base fail-closed evidence aggregator.
- `tools/final-production-gate.ps1` — final fail-closed production aggregator.

Production signing changes the outer installer SHA-256. Therefore exact-installer evidence for production must be generated against the **final signed installer**, not reused from the unsigned development preview.

The final production result must be:

```json
{
  "TestId": "OMNIX-FINAL-PRODUCTION-GATE-002",
  "FailureCount": 0,
  "OverallPass": true
}
```

Anything else is not a production release.

## Building

Local build prerequisites:

- Windows
- Visual Studio 2022 with .NET desktop development
- Visual Studio Tools for Office / Office development build targets
- .NET Framework 4.8 targeting pack
- Inno Setup for installer compilation (the CI workflow provisions it)

Local entry point:

```bat
build\build.bat
```

GitHub Actions builds a **development artifact**. Development preview publication waits for successful exact-head build, architecture-contract, request-budget-runtime and release-provenance-contract workflows. It still does not publish a production release merely because CI is green.

See [docs/CI-VERIFICATION.md](docs/CI-VERIFICATION.md) for the exact CI/human evidence boundary.

## Security and data locations

```text
Settings / protected keys  %LOCALAPPDATA%\OMNIX\settings.dat
Logs                       %LOCALAPPDATA%\OMNIX\logs\
Chat history               %LOCALAPPDATA%\OMNIX\history\
Application                %LOCALAPPDATA%\Programs\OMNIX\
```

Logs and acceptance reports are designed not to contain API keys or full private Office documents. Real release evidence stores hashes/status/timing where possible instead of copying sensitive content.

## Repository layout

```text
OMNIX.sln
src/
  OMNIX.Core/
  OMNIX.Excel/
  OMNIX.Word/
  OMNIX.PowerPoint/
installer/
  installer.iss
build/
  build-with-fallbacks.ps1
  package.ps1
  contract-gates.ps1
  real-evidence-contract-gates.ps1
  final-evidence-contract-gates.ps1
  final-evidence-binding-contract-gates.ps1
  consumer-security-contract-gates.ps1
  sign-production.ps1
tools/
  bound-real-acceptance.ps1
  real-evidence-binding.ps1
  full-office-e2e.ps1
  lifecycle-acceptance.ps1
  consumer-security-acceptance.ps1
  privacy-acceptance.ps1
  release-readiness.ps1
  final-production-gate.ps1
docs/
  CI-VERIFICATION.md
  PRODUCTION-EVIDENCE-RUNBOOK.md
```

## Current honesty boundary

Hosted Windows CI can compile VSTO, validate and hash-bind the payload, build the installer, execute deterministic Core/runtime tests, package exact-head runtime/provenance evidence and scan the development installer. Hosted CI does **not** provide the real desktop Office/restart/consumer-security environment required for production.

A development preview may correctly have `ArtifactType=DEVELOPMENT_ONLY`, `ProductionReleaseApproved=false` and `RealOfficeRuntimeTested=false` even while all required hosted CI workflows pass.

Do not publish the v3 rebuild as production until issue #53 and `OMNIX-FINAL-PRODUCTION-GATE-002` are satisfied on the intended Windows/Office environment.

## License

MIT — see [LICENSE](./LICENSE).

## About

**Powered by Mr Ali**

Creator/support: [Telegram @Ali_silent0](https://t.me/Ali_silent0) · [Telegram channel](https://t.me/Ali_shortcuts) · [Email](mailto:Ali.hekmati2026@gmail.com)

OMNIX is developed as an independent Office/AI integration project. Public contact links should be maintained in the product's About view and repository documentation as current, user-approved project metadata.
