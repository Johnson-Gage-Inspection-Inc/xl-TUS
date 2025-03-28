# Template Repo for handling Excel Files

## Description
This template provides the default configuration for syncing an Excel workbook throughout Johnson Gage & Inspection.

## Configuration

Any time you use this template, you'll need to configure it by setting the variables in the .env file.

By default, the .env file will look something like this:
```
SOP_ID=
SOP_TYPE_ID=
DOC_ID=
DOC_TITLE=
DOC_DETAILS=
```


If your excel file isn't already on Qualer, you'll need to create one [here](https://jgiquality.qualer.com/documents).

Go to the document on qualer, and observe the URL.  It should look something like this:
```
https://jgiquality.qualer.com/Sop/Sop?sopId=2822
```

Copy the `sopId` from the url into your env file as the value for `SOP_ID`.

You may set the other variables however you like, as they just correspond to the Document Title, Document Id, and Document Details as seen in Qualer (The values in the .env file will override the data in Qualer with each PR).
![image](https://github.com/user-attachments/assets/ce6ed96c-2854-4b12-8780-a677724cefcb)

## Structure

```
C:.
│   .env                     # Configuration file, where you'll set the Qualer variables for the particular Excel file
│   .gitattributes           # This tells your system to use git xl for simple diffs, if you have it
│   .gitignore               # This ignores the temporary files that Excel creates when you edit an workbook.
│
└───.github
    └───workflows
            MR-trigger.yml    # This runs the pre-merge checks
            PR-trigger.yml    # This syncs the new version of your excel file to Qualer, 
```

## Workflow

All changes must be sumbitted through a Pull request, which gives an opportunity for automated checks, manual review, and approvial, before merging the proposed changes into the main branch.

### Pull requests

### Automatic checks

Before a pull request can be merged, it must pass a pre-check.  This will fail under certain conditions:

+ If the Excel file was deleted
+ If there are multiple excel files
+ If there are network connectivity issues
+ If the script fails to authenticate with SharePoint (This will happen when an API key expires)
+ If the script fails to authencticate with Qualer (The smallest change to their web interface could cause this, since they don't provide a REST API endpoint for SOP documents)

### Syncing
When a pull request is merged, this will trigger a few things to happen:
+ The file will be imbued with metadata, containing the Release version/tag (The current date) and the git commit hash
+ The file will be uploaded to SharePoint, (If it's an .xltx or .xltm file, it will appear as a template in Excel under the **Johnson Gage & Inspection, Inc** tab when creating a new file.
+ The file will be uploaded to Qualer (with the current date in the filename).
+ The file will be released on GitHub (with the current date in the filename), and release notes will be automatically generated according to the commit history since the last release.

## Diffing
There is a .gitattributes file here, which specificies [git-xl](https://github.com/xltrail/git-xl) as the diffing tool for excel files.
> Warning: git-xl can only diff VBA code, not worksheets or power query.

I recommend xlCompare for anything more in-depth than glancing at diffs in sourcetree.  It's only limitation is that it totally overlooks PowerQuery (M).
