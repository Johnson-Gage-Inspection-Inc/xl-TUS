# TUS Workbook
## Pyrometry

#### 🎯 **Purpose**
The goal of the TUS process is to validate and certify that each industrial oven or furnace meets uniform temperature distribution criteria, as required by **AMS 2750** (aerospace and thermal processing standard). JGI provides certifiable, traceable reports for customers based on detailed sensor data.

---

### 🧪 **Core Workflow**

1. **Setup Phase**
   - The user selects a furnace and customer.
   - Metadata is populated from:
     - `Furnace_Data`
     - `Customer Table`
     - `Salesforce_Data`
   - The oven layout (number of sensors, configuration) is defined by `Interp`.

2. **Data Collection**
   - Thermocouples are installed in specific locations per the furnace layout.
   - The **DAQBook** logs temperature readings during the test cycle.
   - These raw logs are pasted into in `DaqBook_RAW_Data`.

3. **Data Processing & Analysis**
   - **Correction factors** are applied from `Standards_Info` and `Standards_Import`.
   - The corrected data is analyzed in:
     - `Data_Sheet`, `Data_Sheet_15_28`, `Data_Sheet_29_40`
     - `TUS_Worksheet` (for tabulated results)
   - `Comparison_Report` compares this survey to previous ones.

4. **Uniformity Calculation**
   - The system calculates:
     - Min, max, and average temperature
     - Spread/deviation across the survey range
   - These numbers feed into `CERT` (the cert itself).

5. **Certification Packet**
   - A full cert package is compiled using:
     - `CERT`
     - `Packet_Content` (specifies what’s in the packet)
     - `Access_Data` (used for internal database or audit logging)
     - `LOG` (records who did what and when)

6. **Wire Certification Cross-Check**
   - Thermocouple wires have individual Excel certs in a separate folder.
   - A macro reads those files to confirm the wire cert is valid and current.
   - This part of the system is fragile and disorganized, relying on cross-referencing from a directory of Excel files rather than a table.

---

### 🧱 **Oven Layouts and Charts**

- Each **oven** is currently stored in a **separate Excel file with only one row**. These files are meant to hold:
  - Serial number, type, setpoint ranges
  - Sensor locations
  - Any custom deviations or tolerances
- These are compiled dynamically into the main workbook via VBA logic.

**Problem**: This makes reporting, lookup, and updates error-prone and unsustainable.

**Ideal Solution**:
- Replace all those one-row files with a centralized relational table (PostgreSQL or even structured Excel) that the workbook or app can query.
- Each oven would be a row in a `furnaces` table with relationships to sensor positions and layout specs.

---

### 📈 **Visuals & Charts**
- The system uses charts in the `TUS_Worksheet` and possibly in the cert packet to show:
  - Sensor temperature profiles
  - Deviations across time or position
  - Furnace layout visualization (often a box with labeled sensor positions and their results)

---

### 🔐 **Compliance and Logging**
- Everything is tied back to traceable standards and procedures:
  - `LOG` ensures accountability
  - `Software_Validation, Rev. His.` tracks code/release changes
  - `Enabler4Excel_Picklist_Values` and Salesforce links provide standard dropdowns and external traceability

---

### 📚 **Documentation**

For detailed technical information and calculations, see the `docs/` folder:

- **`LiveTUS_Uncertainty_InternalRef.pdf`** - Comprehensive document detailing the uncertainty calculation methodology, statistical analysis, and validation procedures used in the TUS process
- **`LiveTUS_Uncertainty_InternalRef.tex`** - LaTeX source file for `LiveTUS_Uncertainty_InternalRef.pdf`

This document provides the theoretical background and mathematical basis for the temperature uniformity survey calculations and certification process.

---
