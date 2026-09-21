# Office capability status

OMNIX is a Windows VSTO workspace for Excel, Word and PowerPoint. It does not
provide unrestricted computer access, a continuously observed screen, or an
already verified replacement for a database designer/accountant.

## Implemented in this change

- Background instructions identify the active Office host, require inspect/plan/
  approve/apply/read-back, and advertise host-specific write tools.
- `read_document_map` and `read_document_section` navigate the current document
  without selecting cells or opening another document. All calls retain request
  cancellation, document-scope isolation, cloud consent and untrusted-data wrapping.
- Excel: worksheet map in pages of 20; explicit rectangular reads of at most 256
  cells, with values, formulas, number formats and partial-coverage markers.
  Worksheet maps also report direct object-model counts for tables, charts and
  shapes. Per-cell strings are capped at 600 characters. This is not an unbounded
  workbook export.
- Word: `read_document_map` now enumerates available Word story ranges when
  present, including main text, headers/footers, comments, footnotes/endnotes and
  text frames. `read_document_section` reads one selected story in bounded
  character windows of at most 4000. This is object-model inspection, not screen
  capture, and unsupported embedded object internals are not inferred.
- PowerPoint: slide maps report counts for text shapes, tables, groups, pictures
  and charts. Bounded reads can inspect individual shape text, table cells,
  grouped-item metadata and speaker notes. Pixel-only appearance still requires
  slide capture.
- Excel `format_range`: one bounded contiguous range can be professionally formatted after
  explicit preview/confirmation: font name/size/emphasis/color, fill color, horizontal/vertical
  alignment, number format, wrap text, thin/no border, and row/column AutoFit. The actual range is
  brought into view in Excel; formatting is performed through the Excel Object Model.
- Excel `write_to_cell` preserves JSON numbers and booleans as real Excel values while literal
  strings remain text.
- Excel `create_data_table`: one new named worksheet with a styled table, up to
  24 columns/50 data rows and a 32000-character plan. Each write is previewed.
  Existing sheets are never overwritten; `uniqueName=true` can bind a fresh
  suffix before approval. Primitive strings remain literal data. Explicit typed
  cells can store real Excel formulas or ISO dates, with optional number formats.
  Columns are AutoFit with a width cap. Headers, typed cell values/formulas and
  table dimensions are read back before success. Each operation is capped at
  512 cells including headers. A failed operation attempts to remove only its own
  new worksheet and reports cleanup failure rather than claiming atomic success.
  Empty data creates one blank input row. Delete the new sheet to reverse; native
  Ctrl+Z is not guaranteed for this operation.
- Up to 24 provider/tool turns per request, followed by an explicit incomplete-work
  notice if exhausted. This is enough room for ordinary inspect → create → read-back
  multi-sheet workflows while remaining bounded; very large jobs still must be staged.
- Consent/write dialogs and streaming use the pane dispatcher even when the Office
  process has no WPF Application object.

## Vision and provider limits

PNG captures cover only the selected/current view, chart or slide. The selected
API model must accept images. A successful text connection does not demonstrate
image understanding, correct formulas or sound business decisions. No live Gemini
key or Office desktop is available in the hosted build to prove that end to end.

## Still needed for a mature business-system builder

Relational schema/validation and relationships; resumable multi-sheet transaction
plans; formula dependency/error audits; reliable reversible multi-step edits; pivot
tables, chart/report builders; richer Word/PPT creation; arbitrary Ribbon automation;
and actual business acceptance fixtures. Do not describe these as implemented by the
current table/object-model tools.

## Compatibility and evidence

Target matrix: Windows 10/11 with desktop Excel/Word/PowerPoint 2016, 2019, 2021,
2024 and separately identified Microsoft 365 builds, both x86 and x64 Office where
supported by Microsoft. An installation in calendar year 2026 does not itself
identify an Office product/version. Record File > Account > About version/build.
No cell of this full matrix has been certified by this change.

Hosted Windows CI compiles the add-in, runs real compiled WPF startup/dispatcher
checks, plan validation and tool-scope tests. It does not have licensed desktop
Office and cannot validate COM behavior by itself. Use the interactive Office lab
runbook for actual installation, rendering, tool writes, failure rollback, reopening
files, privacy dialogs and UI responsiveness. Windows 10 is already outside normal
OS support; Microsoft documents separate Microsoft 365 security-update provisions.

Sources:
- https://learn.microsoft.com/en-us/visualstudio/vsto/office-solutions-development-overview-vsto
- https://learn.microsoft.com/en-us/microsoft-365-apps/end-of-support/windows-10-support
- https://learn.microsoft.com/en-us/office/ltsc/2024/overview


## Broad Office capability registry

OMNIX now exposes a searchable, explicit capability registry for the active Office host. Models
use `list_office_capabilities` to discover registered operations, `inspect_office_capability`
for bounded read-only inspection, and `apply_office_capability` for one explicit mutation behind
the same preview/confirmation/document-scope boundary as the older focused tools.

The registry substantially expands the programmable Office surface:

- **Excel:** worksheets; bounded ranges; rows/columns; tables; defined names; sorting/filtering;
  data validation; conditional formatting; charts; PivotTable inspection/refresh; cell notes;
  HTTPS/mailto hyperlinks; freeze panes/zoom; page setup; plus the existing formulas, typed values,
  professional formatting and structured table creation.
- **Word:** document/selection/paragraph inspection; styles; tables; bookmarks; fields; comments;
  revisions and Track Changes; content controls; sections/page setup; headers/footers; bounded text,
  font and paragraph formatting; safe allowlisted fields and HTTPS/mailto hyperlinks.
- **PowerPoint:** presentation/slides/sections; shape inspection and geometry; text formatting;
  tables; speaker notes; HTTPS/mailto shape hyperlinks; transitions; basic entrance animations;
  slide layout/move/duplicate/delete and visible slide navigation.

This is intentionally **not arbitrary COM access**. Unknown operation names fail closed. OMNIX does
not expose VBA/macro execution, Trust Center/security changes, arbitrary file/process access,
reflective invocation, or automatic modal-dialog control. Those exclusions are security boundaries,
not missing hidden capabilities.
