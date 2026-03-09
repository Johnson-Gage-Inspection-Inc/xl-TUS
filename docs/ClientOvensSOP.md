[Manage Oven Data and Access Client Spreadsheets](https://app.guidde.com/playbooks/wy3FS7fJJ5NQmjEy3t1pwu)
==========================================================================================================

[Click here to watch](https://app.guidde.com/share/playbooks/wy3FS7fJJ5NQmjEy3t1pwu)

### [![Quick guidde](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FwZAZyxovjT8dAxZkZVPfuT_cover.png?alt=media&token=7ba778ee-ba17-4a93-a328-331c5935c731)](https://app.guidde.com/share/playbooks/wy3FS7fJJ5NQmjEy3t1pwu)

This procedure outlines the steps to add or update oven metadata for TUS calibrations, so critical fields can be populated at calibration time. This is a required step for all new ovens.

### Go to [jgiquality.qualer.com](https://jgiquality.qualer.com)

### 1\. Locate the asset

First, locate the asset on Qualer that you'd like to add data for. We will need it's Asset ID to associate our data with the Qualer asset.

![Locate the asset](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FmxBsUdvCCFpMeAES6dvuiE_doc.png?alt=media&token=1229eaf7-c642-4153-8d57-9cb943c55eea)

### 2\. Access Asset Details

Find the asset and click it's serial number to view its detailed asset information within the Qualer application.

![Access Asset Details](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2Fs7m7CXh9ep2khCH2Cv1edn_doc.png?alt=media&token=2bbf4232-3c8a-49a7-9317-2d37a637cb2c)

### 3\. Open Oven Options pop-out link

Click this pop-out icon to open the asset details in a new tab.

![Open Oven Options pop-out link](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FnoVNA3F8WzBPas17eFyKSY_doc.png?alt=media&token=001f5c8e-1077-49d4-999e-d99445b17300)

### 4\. Copy the assetId from the URL

In the new tab, look at the URL to find the numeric Asset ID for this furnace. Copy it to your clipboard with Ctrl + C.

![Copy the assetId from the URL](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FqWhdPakUGKEwZsuEdu7qrV_doc.png?alt=media&token=89cade2d-e903-490a-b548-c6df4dc6d231)

### 5\. Open ClientOvens.xlsx workbook

With the Asset ID copied, open the ClientOvens Excel workbook. This can be found in the main Pyro folder through OneDrive or Sharepoint

![Open ClientOvens.xlsx workbook](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FpT7TFkBomJd6VVMVA4o9Go_doc.png?alt=media&token=73665414-9dad-443b-93df-082efe772c8b)

### 6\. Add or update a row in the Ovens table, pasting the assetId you copied from Qualer. Fill in the other required fields.

Insert a new row to add a new oven. Paste the AssetId in the first column, and fill out all required fields. If you're unsure which fields are required, ask your immediate supervisor. Existing ovens can be updated here, as well.

![Add or update a row in the Ovens table, pasting the assetId you copied from Qualer. Fill in the other required fields.](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2F94VWrf9npJk3M1tNCcfxWN_doc.png?alt=media&token=d5d6b9fe-aa37-484d-b668-75345da8bd08)

### 7\. Switch from the Ovens sheet to the RangeTol sheet.

Click on the RangeTol tab to switch to the Ranges and Tolerances table.

![Switch from the Ovens sheet to the RangeTol sheet.](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2F5KaZY6y1onv4uGPdhSxjyz_doc.png?alt=media&token=a862ff88-2b71-447e-9150-d84ee7d72887)

### 8\. Add one new row to RangeTol for each range we certify this oven for.

At the bottom of the RangeTol table, paste the AssetId into a new row for each range.

![Add one new row to RangeTol for each range we certify this oven for.](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FpHeWoBjPA4Don89eGQ6TMD_doc.png?alt=media&token=4543f770-3ead-4a58-8921-55651d0d3261)

### 9\. Fill out range and tolerance fields.

For each range, fill in the remaining fields: RangeMin, RangeMax, TolMin, TolMax, TolResolution, and Unit. The TolResolution field is the resolution of the tolerance. Put 0.1 for 1 decimal place, and 1 for no decimals. The other columns can be ignored. For ovens with no minimum range or tolerance, leave RangeMin and TolMin blank, and set TolMax equal to RangeMax.

![Fill out range and tolerance fields.](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2F8TnkpX1dCpbYCQeXyDvdUb_doc.png?alt=media&token=2f97f76f-0451-424c-a471-ffd965fab1a4)

### 10\. Refresh Excel data to propagate the changes

If you have a TUS workbook open already, you may need to click Refresh All in Excel's Data tab to see the updates. New workbooks will grab the latest data automatically.

![Refresh Excel data to propagate the changes](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2F4kkgiQB8weu2N1J5LQ4dEn_doc.png?alt=media&token=77942764-d3e4-4b47-9410-bae94a9ce015)

This procedure ensures that all furnace metadata is available and associated with the correct furnace in Qualer. Unless something changes, this only needs to be done once per furnace.

[Powered by **guidde**](https://www.guidde.com)