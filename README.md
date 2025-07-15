
# On Water Container Report Automation Scripts

This repository contains Python scripts for automating Excel-based **On Water Container Reports**.  
The scripts filter, format, and summarize container shipment data to streamline reporting workflows.

---

## Scripts Overview

### 1. **On Water Report Generator (`B1_on_water_report.py`)**
- Extracts records **without a delivery date** (containers considered "On Water") from the source Excel file.
- Groups and saves the data into a **multi-sheet Excel file** (one sheet per destination).
- Applies header formatting, auto column width, and center alignment.

---

### 2. **Chinese On Water Report Formatter (`B2_chinese_on_water_report.py`)**
- Formats the **On Water Report** by:
  - Inserting **Chinese headers**.
  - Applying borders, fonts, and background colors.
  - Centering text and adjusting column widths.
  - Removing unnecessary columns.
- Outputs a finalized **Chinese-labeled On Water Report**.

---

### 3. **On Water Summary Generator (`B3_on_water_summary.py`)**
- Analyzes the **On Water Report** to count **unique container numbers per destination**.
- Creates a summary Excel file with:
  - Numbering and total container count.
  - Styled title, colored headers, borders, and column adjustments.
  - Grand total summarized at the bottom.

---

### 4. **Departure Report Filter (`B4_depart_report.py`)**
- Filters the **departure data** for a given delivery date range.
- Formats date columns and applies center alignment with adjusted column widths.
- Generates a **filtered departure report** for subsequent analysis.

---

### 5. **Departure Summary Generator (`B5_depart_summary.py`)**
- Summarizes the **filtered departure report** by counting **unique containers per destination**.
- Produces a summary Excel file with:
  - Title and formatted headers.
  - Borders applied only to summary columns.
  - Total container count row added.
  - No formatting beyond specified columns.

---


## Required Python Libraries:
- `pandas`
- `numpy`
- `openpyxl`
- `xlsxwriter`

---

## Notice  
> This is a **backup copy** of the On Water Container Report automation scripts.  
> Ensure to update the file paths, date filters, and other parameters according to working environment.
