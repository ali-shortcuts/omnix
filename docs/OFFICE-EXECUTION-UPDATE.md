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
- **Test model** sends one tiny document-free OMNIX tool-calling probe to the exact selected model.
  A model that can chat but does not return a valid OMNIX tool call is reported as **TextOnly**, not
  as fully working for Office execution.
- **Verify models** tests detected models sequentially with an isolated adapter and per-model timeout,
  classifying OMNIX-tool-compatible, text-only, access-denied, unavailable, rate/quota-limited,
  incompatible, timeout and network failures. The user can stop verification and filter the dropdown
  to models whose OMNIX tool path was actually verified.
- Custom-provider 404 on `/models` is reported as a reachable endpoint with unavailable catalog,
  rather than being mislabeled as failure of the selected model.


## Broad Office capability engine

OMNIX now exposes a queryable, host-specific Object Model capability catalog through
`list_office_capabilities` and a single validated mutation boundary
`execute_office_capability`. The catalog now contains **190 implemented operations**:
**80 Excel**, **61 Word**, and **49 PowerPoint**, in addition to the existing bounded read,
Vision, table-building, formula, formatting, notes, and slide tools.

The capability engine covers major professional surfaces including worksheet/range/data validation, conditional formatting, workbook/worksheet protection, calculation, outline/grouping,
filters, sorting, names, tables, charts and page layout in Excel; editing, font/paragraph styles,
lists, tables, headers/footers, bookmarks, comments, review/revisions, references and page layout
in Word; and slides, shapes, text, tables, arrange/z-order, links, notes, animation and transitions
in PowerPoint.

This is deliberately **not** an unrestricted Office escape hatch. Arbitrary `ExecuteMso`, VBA/macro
execution, Trust Center/security changes, shell/registry access, and arbitrary file-system access
remain unavailable. Every catalogued mutation still goes through document-scope validation,
preview, explicit user confirmation, real Office Object Model execution, and visible target reveal.


## Native provider tool runtime

OMNIX no longer relies only on models reproducing a textual `omnix_tool` block correctly.

- OpenAI-compatible providers receive a real `tools` schema for a single `omnix_tool` function.
- Anthropic-compatible custom endpoints receive an Anthropic `tools/input_schema` definition.
- Gemini receives `functionDeclarations` and native `functionCall` responses are parsed.
- Ollama receives a native tool definition when the local model/runtime supports tools.
- If a compatible endpoint explicitly rejects the native tool-schema fields with a request-shape
  error, OMNIX retries that provider turn once using the legacy textual protocol. The whitelist,
  document scope and write confirmation boundaries do not change.
- Native/text tool names are conservatively normalized only when they resolve to an existing hard
  whitelist entry. Multiple parallel calls, malformed argument objects and invented tools fail
  closed and receive bounded repair turns.
- Explicit create/edit requests run an authoritative Office preflight. A model may not finish a
  mutation request as prose-only output before attempting a real write. If bounded repair fails,
  OMNIX reports a runtime failure instead of fabricating a table, VBA macro or success claim.
- Gateway diagnostics record response kind, sanitized tool name, whitelist/write classification and
  success/failure counts. Tool arguments, API keys and Office document content are not logged.
- A successful write is no longer sufficient for a success answer. The latest Office mutation is marked
  unverified until a subsequent real Office read tool reads back document state. If a provider tries to
  finish before read-back, the Gateway issues bounded verification-repair turns; repeated failure returns
  an explicit incomplete/runtime result instead of claiming that the workbook/document/presentation is done.
- Legacy textual fallback tool names may include common provider namespaces (for example
  `omnix.write_to_cell`); normalization is accepted only when it resolves to an existing hard-whitelisted
  canonical tool. This improves compatibility without expanding the executable surface.
