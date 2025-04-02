# Import VBA Subroutines Documentation

## Overview

This file contains two VBA subroutines, `ImportFiles` and `ImportHRM`, which are used to import data from external files into an Excel workbook. These subroutines handle the extraction of data from SAP and HRM systems and organize it into specific worksheets for further processing.

---

## Subroutine: `ImportFiles`

### Purpose

The `ImportFiles` subroutine imports data from an Excel file containing picking and replenishment information. It copies the relevant data into the active workbook and organizes it into a worksheet named `P&R Lines`.

---

### Steps

1. **Prompt for File Selection**:

   - The user is prompted to select an Excel file containing picking and replenishment data.

2. **Open the Selected File**:

   - If a file is selected, it is opened as a new workbook (`wbPickedLines`).

3. **Copy Data**:

   - Each visible sheet in the selected workbook is copied into the active workbook.
   - The existing `P&R Lines` sheet in the active workbook is deleted.
   - The copied sheet is renamed to `P&R Lines`.

4. **Close the Opened Workbook**:

   - The selected workbook is closed after the data is copied.

5. **Set Variables**:

   - The `P&R Lines` sheet is assigned to a variable (`wPL`).

6. **Select the `Data` Sheet**:
   - The `Data` sheet in the active workbook is selected.

---

### Inputs

- **File Selection**:
  - The user is prompted to select an Excel file (`*.xls` or `*.xlsx`) containing picking and replenishment data.

---

### Outputs

- **Updated Workbook**:
  - The `P&R Lines` sheet in the active workbook is updated with data from the selected file.

---

### Error Handling

- If no file is selected, a message box is displayed, and the subroutine exits.

---

## Subroutine: `ImportHRM`

### Purpose

The `ImportHRM` subroutine imports HRM data from a text file into the active workbook. It creates a new worksheet named `HRM` and populates it with the imported data.

---

### Steps

1. **Prompt for File Selection**:

   - The user is prompted to select a text file containing HRM data.

2. **Create the `HRM` Sheet**:

   - If a file is selected, the existing `HRM` sheet is deleted (if it exists).
   - A new sheet is created and named `HRM`.

3. **Import Data**:

   - The selected text file is imported into the `HRM` sheet using a query table.

4. **Set Default Values**:

   - The first row of the `HRM` sheet is set to "N".

5. **Select the `Data` Sheet**:
   - The `Data` sheet in the active workbook is selected.

---

### Inputs

- **File Selection**:
  - The user is prompted to select a text file (`*.txt`) containing HRM data.

---

### Outputs

- **Updated Workbook**:
  - A new `HRM` sheet is created in the active workbook, populated with data from the selected text file.

---

### Error Handling

- If no file is selected, a message box is displayed, and the subroutine exits.

---

## Key Components

### Subroutine: `ImportFiles`

| **Step**               | **Description**                                                                      |
| ---------------------- | ------------------------------------------------------------------------------------ |
| File Selection         | Prompts the user to select an Excel file containing picking and replenishment data.  |
| Open Workbook          | Opens the selected file as a new workbook.                                           |
| Copy Data              | Copies visible sheets into the active workbook and renames the sheet to `P&R Lines`. |
| Close Workbook         | Closes the selected workbook after copying data.                                     |
| Update Active Workbook | Updates the `P&R Lines` sheet in the active workbook.                                |

---

### Subroutine: `ImportHRM`

| **Step**               | **Description**                                                        |
| ---------------------- | ---------------------------------------------------------------------- |
| File Selection         | Prompts the user to select a text file containing HRM data.            |
| Create `HRM` Sheet     | Deletes the existing `HRM` sheet (if it exists) and creates a new one. |
| Import Data            | Imports data from the selected text file into the `HRM` sheet.         |
| Set Default Values     | Sets the first row of the `HRM` sheet to "N".                          |
| Update Active Workbook | Updates the workbook with the new `HRM` sheet.                         |

---

## Assumptions

1. The user selects valid files for importing data:
   - An Excel file (`*.xls` or `*.xlsx`) for `ImportFiles`.
   - A text file (`*.txt`) for `ImportHRM`.
2. The `P&R Lines` and `HRM` sheets exist in the workbook or can be created as needed.
3. The imported data is structured correctly for further processing.

---

## Limitations

1. **Hardcoded Sheet Names**:

   - The subroutines rely on specific sheet names (`P&R Lines`, `HRM`, and `Data`). Changes to these names will require updates to the code.

2. **No Validation**:

   - The subroutines do not validate the structure or content of the selected files.

3. **Error Handling**:
   - The subroutines use `On Error Resume Next`, which may suppress critical errors and make debugging difficult.

---

## Example Usage

### ImportFiles

1. Run the `ImportFiles` subroutine.
2. Select an Excel file containing picking and replenishment data.
3. Verify that the `P&R Lines` sheet in the active workbook is updated with the imported data.

### ImportHRM

1. Run the `ImportHRM` subroutine.
2. Select a text file containing HRM data.
3. Verify that the `HRM` sheet in the active workbook is created and populated with the imported data.

---

## Recommendations for Improvement

1. **Error Handling**:

   - Add robust error handling to manage issues such as invalid file formats or missing sheets.
   - Example:
     ```vba
     On Error GoTo ErrorHandler
     ' Code here
     Exit Sub
     ErrorHandler:
     MsgBox "An error occurred: " & Err.Description
     ```

2. **Dynamic Sheet Names**:

   - Allow the user to specify sheet names dynamically instead of hardcoding them.

3. **Validation**:
   - Validate the structure and content of the selected files before importing data.

---
