# Office execution and workspace update

## Changed

- Parse fenced and `<tool_call>omnix_tool ...</tool_call>` calls; malformed or multiple calls fail closed. The existing whitelist, cancellation, scope and write-confirmation boundaries remain enforced.
- Hide both streaming protocol markers. Report real tool start/result, without simulating ribbon clicks.
- Multiline input with up to 200px height and scrolling. Reject oversized drafts before clearing them; ignore Send while busy. Full current messages up to 65,536 characters are sent; this is not unlimited model context.
- UTF-8 TXT import within the same text limit, Copy all, and a selectable Transcript view (Ctrl+A/Ctrl+C). Individual bubbles remain separate selectable controls.
- Prioritize up to 20 recent turns within the existing request budget. A read-only conversation-search tool retrieves excerpts of earlier loaded turns for this document. Stored-history retention still applies.
- Custom Provider first; local providers last. Explicit Custom Model entry for every provider, existing editable model IDs preserved. General/privacy controls are collapsed separately; no global privacy protections removed.
- Save reports storage failure and returns to Chat only after successful persistence.
- Learn page with separate Excel/Word/PowerPoint selections and model-accessible reference search.

- Excel cell writes accept an explicit worksheet name in the active workbook. Formula writes use General format and verify HasFormula; successful table creation leaves the new sheet visible and wraps headers. These COM changes require real Excel validation beyond hosted tests.

## Reference provenance

The Excel index contains 521 distinct names from 514 nonempty linked entries in Microsoft's alphabetical index, retrieved 2026-09-21. Combined labels (such as FIND/FINDB) are separated; the index label typo BETA.INVn is normalized to BETA.INV. The count is a documentation snapshot, not a promise that every version supports every function. Only names and official links are embedded; no Microsoft descriptions are copied wholesale.

- https://support.microsoft.com/en-us/excel/excel-functions-alphabetical
- https://support.microsoft.com/en-us/word/use-a-formula-in-a-word-table

Word formulas are fields and are not the Excel function engine. PowerPoint is not a worksheet-function engine. The reference does not grant new Office write abilities.

## Limits and validation

This update does not implement universal ribbon automation or unrestricted full-document vision. Existing bounded structural reads and captures remain scoped to the active document; headers, comments, embedded objects and other unsupported areas must not be represented as fully visible. Existing write tools remain the supported execution surface. Agent Router credentials and every remote model cannot be validated without the user's configured service.

Windows build/runtime acceptance must pass before publishing an installer. GitHub-hosted checks do not substitute for installed Excel/Word/PowerPoint tests across Office versions. UI XML and whitespace checks are only preliminary validation, not runtime proof.


## Live workspace and provider-tool compatibility

- Chat bubbles now resolve/copy theme brushes into the FlowDocument itself so dark-mode
  text does not fall back to black when Office hosts WPF without a normal Application
  resource tree.
- The tool-call parser accepts the existing OMNIX fenced/XML protocol plus provider-native
  textual calls such as `<|tool_call_start|>[write_to_cell(...)]<|tool_call_end|>`.
  Parsing remains data-only: no reflection/eval is used, only one call is accepted, and
  the normal whitelist, scope guard and write confirmation still apply.
- Provider-native tool protocol markers are filtered from streaming UI output.
- Chat no longer contains a Live activity timeline. Execution visibility belongs in the Office
  application itself.
- Excel/Word/PowerPoint host adapters implement visible execution: OMNIX activates the relevant
  native Ribbon tab when possible and brings the real worksheet/range, Word range, or slide/shape
  into view before/after the actual operation. Object-model writes are never misrepresented as
  fake Ribbon-button clicks.
- The system prompt is rebuilt from live Office context before every provider/tool turn, so after
  a sheet/selection/slide change the model is explicitly reminded that it is operating through
  OMNIX inside the current Excel/Word/PowerPoint document.
- Learn now has a local English/Persian selector. This changes the reference presentation
  only; installer and provider/settings behavior are unchanged.


## Provider/model diagnostics separation

Settings now treats provider connectivity and model inference as different facts:

- **Test connection** performs a model-independent authenticated catalog/endpoint check. It never
  sends a chat request and never fails merely because the selected model is unavailable.
- **Detect models** retrieves the provider-advertised catalog only; discovery is not presented as
  proof that inference works.
- **Test model** sends one tiny document-free synthetic text request to the exact selected model.
- **Verify models** tests detected models sequentially with an isolated adapter and per-model timeout,
  classifying working, access-denied, unavailable, rate/quota-limited, incompatible, timeout and
  network failures. The user can stop the verification and filter the dropdown to verified-working
  catalog entries.
- Custom-provider 404 on `/models` is reported as a reachable endpoint with unavailable catalog,
  rather than being mislabeled as failure of the selected model.

