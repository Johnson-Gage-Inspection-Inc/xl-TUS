This repository tracks an excel workbook for performing Termperature Uniformity Surveys.  Since GitHub isn't well-equiped to diff excel files, I'm leveraging a few scripts to explode the parts of the workbook we want to track into plain text for easier diffing and review.

- ExportPowerQuery.bas - exports all PowerQuery scripts to .m files
- ExportVBA.bas - exports all VBA scripts (including itself and ExportPowerQuery.bas) to .bas or .cls files
- explode_excel.py - Exports each worksheet as a .csv file, prefering formulas over values.

So, as you're commenting on PR's, please refer to the things these represent (So, refer to `Main` (the sheet) rather than `Main.csv` (the file) when summarizing or suggesting changes.
