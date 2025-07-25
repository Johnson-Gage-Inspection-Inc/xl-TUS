# _Temperature Uniformity Survey_ Workbook
> Pyrometry

This is the official repository for tracking and maintaining the Excel workbook for TUS certificates.  This repository houses [Usage instructions](#usage-instructions), [releases](https://github.com/Johnson-Gage-Inspection-Inc/xl-TUS/releases), [changes](https://github.com/Johnson-Gage-Inspection-Inc/xl-TUS/compare), and other [Documentation](#documentation).

<!-- TOC tocDepth:2..3 chapterDepth:2..6 -->

- [Usage instructions](#usage-instructions)
   - [Initialization (One time per user, per machine)](#initialization-one-time-per-user-per-machine)
      - [Trust settings](#trust-settings)
      - [Sharepoint Authentication](#sharepoint-authentication)
      - [JGI server authentication](#jgi-server-authentication)
   - [**Core Workflow** (Once per calibration)](#core-workflow-once-per-calibration)
- [**Documentation**](#documentation)

<!-- /TOC -->
---

## Usage instructions

### Initialization (One time per user, per machine)

#### Trust settings

1. When you first open the template, you may have to trust the document by clicking `Yes` on this window.

   ![image](https://github.com/user-attachments/assets/b04805df-24e9-457a-b5fc-4834d254861e)

#### Sharepoint Authentication

2. Soon, you'll be asked to authenticate to sharepoint.com. Select `Organizational account` and click `Sign in` to be redirected to an OAuth login window.

   ![image](https://github.com/user-attachments/assets/13446558-20d6-4e32-9b2b-14473747c75c)


3. Select your work account (you@jgiquality.com)

   ![image](https://github.com/user-attachments/assets/a09f10c1-6549-4cf8-857e-487315c2a98a)

4. Now, you should see "You are currently signed in." Click `Connect`

   ![image](https://github.com/user-attachments/assets/bb8451b9-8fec-42b7-981b-f288ca042c97)

#### JGI server authentication

6. When prompted for `https://jgiapi.com`, just click `Connect`.

   ![image](https://github.com/user-attachments/assets/f9cd3763-809f-4b31-9823-90032e1669d8)

   > This will return a warning "We couldn't authenticate with the credentials provided. Please try again". This is expected.

   ![image](https://github.com/user-attachments/assets/b59042bd-ea69-4f31-97d7-525936571f6c)

7. Now, you'll by asked to authenticate to `https://jgiapi.com/wire-offsets`.  Repeat steps 2-4 for this endpoint.

   ![image](https://github.com/user-attachments/assets/0b0a41b2-5ece-40a6-aa43-c30a7775962d)

   ![image](https://github.com/user-attachments/assets/29da3139-63ea-459f-b5dd-39557ae92e1b)


### **Core Workflow** (Once per calibration)

1. **Setup Phase** (Requires internet connection)
   - Initial user inputs (Requires internet. For offline use, do this first, then save the file)
    - Work order number (i.e. 56561-012345, or simply 12345)
    - Work item number (defaults to 1)
    - Daqbook asset tag (i.e. J1)
   > Metadata for the order, customer, and UUT and offsets for reference standards will be automatically populated.

2. For offline use, save the file to your flashdrive (or laptop?) before you leave JGI (or otherwise disconnect from the internet).  This will save the offsets and metadata within your workbook, for later use.

   ![image](https://github.com/user-attachments/assets/af2ede84-db30-4cd6-b4bd-b1cb21327d2f)

   Make sure to save it as a `.xlsm` file (Macro enabled),  **NOT** .xlsx (default).

   ![image](https://github.com/user-attachments/assets/c095c94a-bb21-43f4-825b-91c63ce32dfd)

3. Once the file is saved, you're free to close excel and proceed to the work site.

## **Documentation**

For detailed technical information and calculations, see the `docs/` folder:

- [**`LiveTUS_Uncertainty_InternalRef.pdf`**](https://github.com/Johnson-Gage-Inspection-Inc/xl-TUS/blob/main/docs/LiveTUS_Uncertainty_InternalRef.pdf) - Comprehensive document detailing the uncertainty calculation methodology, statistical analysis, and validation procedures used in the TUS process. This document provides the theoretical background and mathematical basis for the temperature uniformity survey calculations and certification process.

---
