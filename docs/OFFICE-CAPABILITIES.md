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
  cells, with values/formulas and partial-coverage markers. Per-cell strings are
  capped at 600 characters. This is not an unbounded workbook export.
- Word: main-story text in character windows of at most 4000. Headers, footers,
  footnotes, comments and text boxes are not included by this new read tool.
- PowerPoint: slide map in pages of 20 and paginated individual shape text.
  Grouped objects, tables and notes are not fully covered by the new text reader.
  Existing slide capture remains useful for visible non-text content.
- Excel `create_data_table`: one new named worksheet with a styled table, up to
  24 columns/50 data rows and a 32000-character plan. Each write is previewed.
  Existing sheets are not overwritten. Strings are literal, not executed formulas.
  Headers, cell values and table dimensions are read back. Each operation is capped at 512 cells including headers. A failed operation attempts to remove
  only its own new worksheet and reports cleanup failure rather than claiming
  atomic success. Empty data creates one blank input row. Delete the new sheet
  to reverse; native Ctrl+Z is not guaranteed for this operation.
- Eight provider turns per request, followed by an explicit incomplete-work notice
  if exhausted. Large jobs must be staged; this does not promise autonomous
  completion of an arbitrarily large workbook.
- Consent/write dialogs and streaming use the pane dispatcher even when the Office
  process has no WPF Application object.

## Vision and provider limits

PNG captures cover only the selected/current view, chart or slide. The selected
API model must accept images. A successful text connection does not demonstrate
image understanding, correct formulas or sound business decisions. No live Gemini
key or Office desktop is available in the hosted build to prove that end to end.

## Still needed for a mature business-system builder

Typed schema/validation and relationships; multi-sheet plans with resumable progress;
formula dependency/error checks; reliable reversible multi-step edits; pivot tables,
charts and reports; richer Word/PPT creation and coverage; actual business acceptance
fixtures. Do not describe these as implemented by the new table tool.

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
