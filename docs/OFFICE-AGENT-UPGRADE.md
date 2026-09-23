# Office agent upgrade — 2026-09-23

## Evidence and implemented changes

User runtime logs show a 711 ms table write followed by a 210,467 ms provider round. These timings do not alone prove the cause of a UI freeze. Code inspection found one dispatcher callback and full Markdown reconstruction for every stream delta. This upgrade coalesces deltas every 150 ms at background dispatcher priority and incrementally displays plain text, rendering final Markdown once. Provider transport/retries run outside Office's STA; Office context and mutations stay on their owning dispatcher. Each provider round has a 120-second cancellation deadline. Applied writes are retained if later inference fails.

Verified model results are selectable and include persistent multiple saved IDs per provider; one model remains active. Text-only models are distinguished from tool-compatible models. API key, endpoint/protocol and Cloudflare account changes invalidate verification. Model IDs remain case-sensitive. Manual IDs and saved choices survive catalog refresh.

New sheets can have `title` and `startRow` arguments. A title occupies a framed, merged heading above the table; its default header row is 4 and first data row is 5. Formula references must match that location. Existing content is not overwritten. `sheet.heading` adds a heading only to an empty bounded range, and verifies its text. Existing table creation still verifies typed values, formulas, headers and dimensions. Formula-looking primitive strings are rejected with corrective guidance instead of silently displaying unevaluated formulas; intentional literal text can be prefixed with an apostrophe. Formulas must use explicit typed objects.

Host-specific system guidance now covers inspect/plan/execute/read-back/repair, sensible simple-prompt defaults, continuing on existing sheets, separate titles, units, typed formulas, layout review and honest completion. Business locale defaults to Afghanistan/Dari/AFN and is editable. Instructions explicitly defer to the user's stated choices and do not invent real business records.

Optional visible pacing yields the UI before each real Office write and rechecks cancellation/document identity. It does not animate fictitious Ribbon clicks. Routine new-table creation and formatting can proceed without repeated approval; other mutations keep confirmation. “Confirm every edit” restores all-write confirmation. CloudAllowed is the fresh-install default; saved privacy choices remain unchanged.

## Provider sources

- SiliconFlow: `https://api.siliconflow.com/v1`; Bearer auth; discovery `GET /models?sub_type=chat`. Selected models are free, others are paid. https://docs.siliconflow.com/en/userguide/quickstart and https://docs.siliconflow.com/en/api-reference/models/get-model-list
- Cloudflare: account-specific `/client/v4/accounts/{account_id}/ai/v1`; Bearer API token. Discovery uses `/ai/models/search?format=openrouter&task=Text%20Generation&per_page=100&page=...`, not an assumed `/v1/models`. Daily free allocation applies to eligible models; some need paid plans. https://developers.cloudflare.com/workers-ai/configuration/open-ai-compatibility/ and https://developers.cloudflare.com/api/resources/ai/subresources/models/methods/list/ and https://developers.cloudflare.com/workers-ai/platform/pricing/
- New presets do not claim verified vision. Native tool compatibility still requires the existing synthetic model probe. No automatic paid fallback was added.

## Requirements audit

| User request | Status / evidence |
|---|---|
| Freeze while replying, Stop and long waits | Coalesced incremental streaming, background provider transport, bounded rounds; Windows dispatcher heartbeat regression |
| Custom first; concise names; OpenAI/Anthropic; manual models; correct Save navigation | Existing paths retained; Windows UI and both-protocol transport regressions |
| Click verified model; save several | Added selectable complete tested-result list and saved-model collection |
| Provider-specific fields and local controls | Retained; Cloudflare account field added only for Cloudflare |
| Fewer cloud/format dialogs | Fresh-install CloudAllowed, configurable routine-write auto-apply; explicit saved privacy preserved |
| Professional title, no duplicate follow-up sheets, real formulas | Native heading and titled-table layout plus explicit workflow guidance; duplicate prevention remains enforced by tools; model reasoning still needs quality evaluation |
| Know current host/document/selection and actual tools | Existing live context, access probe and queryable host capability catalog retained and used in guidance |
| Real-time visible execution | Existing real selection/Ribbon reveal plus yielding/pacing; no claim of support for every Ribbon command |
| Agent loop and result checking | Existing bounded tool loop/read-back recovery retained; professional playbooks strengthened |
| Long input, scrolling, TXT, transcript and history | Existing implementations retained; context/transport remain bounded, not unlimited |
| Learn/formula reference separately per Office app | Existing Learn/reference search retained; Word/PPT are not misrepresented as Excel function catalogs |
| More free providers | SiliconFlow and Cloudflare added with official routes and qualified pricing metadata |
| Hundreds of curated visual templates | Not delivered; current blueprints are guidance, not hundreds of tested design assets |
| Full-document vision and every Office tool | Not an honest universal guarantee: structured reads are paged; captures show bounded views; model vision support and Office permissions still apply |
| Windows 10/11, Office 2016+ | Installer architecture retained; actual host/version acceptance matrix remains required |
| Crash recovery and no false success | Existing encrypted history and read-back checks retained; logs cannot prove Excel COM stability across every installation |
| Remove all other versions | No unmerged user branches removed. Replace released installer only after successful build/provenance checks |

## Validation and manual acceptance

Windows CI covers compilation, WPF initialization/selection, privacy, encrypted history, provider diagnostics, request budgets, protocol fixtures and installer packaging. New tests exercise heading plan layout, selectable verified models, dispatcher heartbeat during deliberately blocking provider setup and the Cloudflare model discovery route.

GitHub-hosted Windows does not include a usable installed Excel/Word/PowerPoint environment. Manual acceptance must confirm: a titled gold-shop workbook using AFN; numeric/formula read-back; edits reuse the correct sheet; Stop during streaming and before mutation; no Excel restart; small-pane RTL readability; Cloudflare/SiliconFlow credentials and quota behavior; supported Office x86/x64 versions. Passing CI is not evidence of zero defects or production Authenticode signing.
