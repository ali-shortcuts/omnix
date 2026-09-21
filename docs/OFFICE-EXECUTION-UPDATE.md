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
