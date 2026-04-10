---
name: tus-workbook-git-tracking
description: 'Track Excel workbook changes for TUS.xltm in Git by exporting VBA, Power Query M, and worksheet formulas. Use when updating macros, queries, names, or sheet logic and preparing reviewable diffs.'
argument-hint: 'Workbook path (for example: TUS.xltm or examples/J202172501.xlsm)'
---

# TUS Workbook Git Tracking

## Purpose
Convert workbook internals into stable text artifacts so code review can focus on meaningful behavior changes instead of opaque binary differences.

## Use When
- You changed VBA code in the workbook.
- You changed Power Query scripts or query connection behavior.
- You changed worksheet formulas, names, or table structures.
- You are preparing a pull request and need reviewable diffs.

## Inputs
- Workbook file, usually TUS.xltm.
- Export macro modules in the workbook project:
  - exploded/TUS/macros/ExportVBA.bas
  - exploded/TUS/macros/ExportPowerQuery.bas
- Python export script:
  - explode_excel.py

## Workflow
1. Open the target workbook in Excel.
2. Run ExportVisualBasicCode from ExportVBA.bas.
Outcome: VBA components are exported to exploded/<WorkbookName>/macros as .bas/.cls/.frm and converted to UTF-8 without BOM.
3. Run ExportAllQueryMCode from ExportPowerQuery.bas.
Outcome: query .m files are exported to exploded/<WorkbookName>/queries with Authorization tokens redacted and connection properties included.
4. Run explode_excel.py with the workbook path.
Outcome: formula-first sheet snapshots are exported to exploded/<WorkbookName>/sheets and Name Manager entries to exploded/<WorkbookName>/names.tsv.
5. Review the resulting text diffs before commit.

## Decision Points
- If only VBA changed:
  - Run step 2 and review exploded/<WorkbookName>/macros.
- If only Power Query changed:
  - Run step 3 and review exploded/<WorkbookName>/queries.
- If sheet formulas, tables, or names changed:
  - Run step 4 and review exploded/<WorkbookName>/sheets plus names.tsv.
- If workbook-wide refactors happened:
  - Run all steps 2 through 4 to keep artifacts synchronized.

## Quality Checks
- All expected files exist under exploded/<WorkbookName>/macros, queries, and sheets.
- No secret tokens appear in exported .m files.
- Exported text is UTF-8 and diffable.
- names.tsv is updated when defined names or tables changed.
- Diff descriptions use workbook semantics in review language:
  - Refer to Main, Data_Sheet, and other sheets by sheet name.
  - Avoid framing findings as CSV file behavior.

## Completion Criteria
- Binary workbook change is accompanied by matching text artifact updates.
- Diffs are reviewable and map clearly to VBA, Power Query, and worksheet behavior.
- Pull request summary explains behavioral impact at workbook or sheet level.

## Example Prompts
- "Track changes in TUS.xltm and tell me what changed in VBA, queries, and sheets."
- "I changed formulas on Main and Data_Sheet. Run the workbook export workflow and summarize review-ready diffs."
- "Prepare artifact updates for this .xlsm and check for redaction or encoding problems."
