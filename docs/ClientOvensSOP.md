[Manage Oven Data and Access Client Spreadsheets](https://app.guidde.com/playbooks/wy3FS7fJJ5NQmjEy3t1pwu)
==========================================================================================================

[Click here to watch](https://app.guidde.com/share/playbooks/wy3FS7fJJ5NQmjEy3t1pwu)

### [![Quick guidde](https://static.guidde.com/v0/qg%2FQ75mVVtuOoMqsupZtgCRhOCKwDi2%2Fwy3FS7fJJ5NQmjEy3t1pwu%2FwZAZyxovjT8dAxZkZVPfuT_cover.png?alt=media&token=7ba778ee-ba17-4a93-a328-331c5935c731)](https://app.guidde.com/share/playbooks/wy3FS7fJJ5NQmjEy3t1pwu)

This procedure outlines the steps to add or update oven metadata for TUS calibrations, so critical fields can be populated at calibration time. This is a required step for all new ovens.

### Log into [jgiquality.qualer.com](https://jgiquality.qualer.com)

### 1\. Locate the asset

First, locate the asset on Qualer that you'd like to add data for. We will need it's Asset ID to associate our data with the Qualer asset.

<img width="1920" height="944" alt="image" src="https://github.com/user-attachments/assets/f2e21c6a-8a53-45ec-86b1-76b83b447c33" />

### 2\. Access Asset Details

Find the asset and click it's serial number to view its detailed asset information within the Qualer application.

<img width="1920" height="945" alt="image" src="https://github.com/user-attachments/assets/ee445362-64f5-4339-aabc-f419bb6c3ec6" />

### 3\. Open Oven Options pop-out link

Click this pop-out icon to open the asset details in a new tab.

<img width="1920" height="945" alt="image" src="https://github.com/user-attachments/assets/14782f52-241a-4553-9237-eec033c329eb" />

### 4\. Copy the assetId from the URL

In the new tab, look at the URL to find the numeric Asset ID for this furnace. Copy it to your clipboard with Ctrl + C.

<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/264e3dce-7b95-4185-a9db-8d72c01c4e64" />

### 5\. Open ClientOvens.xlsx workbook

With the Asset ID copied, open the ClientOvens Excel workbook. This can be found in the main Pyro folder through OneDrive or Sharepoint

<img width="1920" height="945" alt="image" src="https://github.com/user-attachments/assets/5be24f78-b784-4d99-aa11-100359872eb7" />

### 6\. Add or update a row in the Ovens table, pasting the assetId you copied from Qualer. Fill in the other required fields.

Insert a new row to add a new oven. Paste the AssetId in the first column, and fill out all required fields. If you're unsure which fields are required, ask your immediate supervisor. Existing ovens can be updated here, as well.

<img width="1920" height="945" alt="image" src="https://github.com/user-attachments/assets/073d6309-e267-4bcd-afa8-758f3ca16fa8" />

### 7\. Switch from the Ovens sheet to the RangeTol sheet.

Click on the RangeTol tab to switch to the Ranges and Tolerances table.

<img width="1920" height="945" alt="image" src="https://github.com/user-attachments/assets/29c4ecc9-b6c8-4b91-9805-2d5667335b8a" />

### 8\. Add one new row to RangeTol for each range we certify this oven for.

At the bottom of the RangeTol table, paste the AssetId into a new row for each range.

<img width="1920" height="945" alt="image" src="https://github.com/user-attachments/assets/523218e9-781d-4374-9184-4ab870da4234" />

### 9\. Fill out range and tolerance fields.

For each range, fill in the remaining fields: RangeMin, RangeMax, TolMin, TolMax, TolResolution, and Unit. The remaining columns can be ignored.
The TolResolution field is the resolution of the tolerance. Put 0.1 for 1 decimal place, and 1 for no decimals.
For ovens with no minimum range or tolerance, leave RangeMin and TolMin blank, and set TolMax equal to RangeMax.

<img width="1920" height="998" alt="image" src="https://github.com/user-attachments/assets/ce16375a-5a9e-48e4-bcc1-3949c29f559f" />

### 10\. Refresh Excel data to propagate the changes

If you have a TUS workbook open already, you may need to click Refresh All in Excel's Data tab to see the updates. New workbooks will grab the latest data automatically.

<img width="1920" height="1030" alt="image" src="https://github.com/user-attachments/assets/e3604005-40d7-4cbc-bae7-697835fafc67" />

This procedure ensures that all furnace metadata is available and associated with the correct furnace in Qualer. Unless something changes, this only needs to be done once per furnace.

[Powered by **guidde**](https://www.guidde.com)
