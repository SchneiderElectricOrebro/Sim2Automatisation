# Run VBA Module Documentation

## Overview

This file contains two main subroutines, `Run` and `Consolidation`, along with several public variables. The `Run` subroutine executes daily logic based on the current day of the week, while the `Consolidation` subroutine performs data aggregation and calculations for various tasks. These subroutines are designed to automate data processing in an Excel workbook.

---

## Public Variables

### Purpose

The public variables are used to store intermediate data and counters for various tasks, such as `iOrdTruck`, `iHighLift`, `iSmalGang`, etc. These variables are shared across the subroutines.

### List of Variables

| **Variable Name**               | **Type**    | **Description**                                       |
| ------------------------------- | ----------- | ----------------------------------------------------- |
| `iRow`                          | `Long`      | Row index for iterating through data.                 |
| `LastRowPL`                     | `Range`     | Stores the last row in the "P&R Lines" sheet.         |
| `wPL`                           | `Worksheet` | Reference to the "P&R Lines" worksheet.               |
| `iOrdTruck`                     | `Long`      | Counter for "Ordinary Truck" tasks.                   |
| `iHighLift`                     | `Long`      | Counter for "High Lift" tasks.                        |
| `iSmalGang`                     | `Long`      | Counter for "Small Gang" tasks.                       |
| `iLongGoods`                    | `Long`      | Counter for "Long Goods" tasks.                       |
| `iPaternost`                    | `Long`      | Counter for "Paternoster" tasks.                      |
| `iRepl`                         | `Long`      | Counter for replenishment tasks.                      |
| `iPackA1`, `iPackA2`, `iPackA3` | `Double`    | Counters for packing tasks for shifts A1, A2, and A3. |
| `iInbo`                         | `Long`      | Counter for inbound tasks.                            |

---

## Subroutine: `Run`

### Purpose

The `Run` subroutine executes daily logic based on the current day of the week. It clears previous data, applies logic for the current day, and calls specific subroutines for each weekday.

---

### Steps

1. **Initialize Variables**:

   - Resets all public variables to `0`.

2. **Clear Previous Data**:

   - Clears specific ranges in the "Data" sheet.
   - Removes any active filters in the "HRM" and "P&R Lines" sheets.

3. **Determine Current Day**:

   - Uses the `Format(Now(), "DDD")` function to determine the current day of the week.

4. **Call Day-Specific Subroutines**:
   - Calls the appropriate subroutine (`Monday`, `Tuesday`, etc.) based on the current day.

---

### Inputs

- **Current Date**:
  - The subroutine uses the current date to determine the day of the week.

---

### Outputs

- **Updated Workbook**:
  - Clears and updates data in the "Data" sheet based on the logic for the current day.

---

### Limitations

1. **Hardcoded Day Logic**:
   - The subroutine relies on specific subroutines (`Monday`, `Tuesday`, etc.) that must be defined elsewhere in the project.
2. **No Error Handling**:
   - The subroutine does not handle errors, such as missing sheets or invalid data.

---

## Subroutine: `Consolidation`

### Purpose

The `Consolidation` subroutine performs data aggregation and calculations for various tasks, such as HRM data, picking and replenishment, packing, and inbound tasks. It updates the "Data" sheet with the results.

---

### Steps

1. **HRM Data**:

   - Applies formulas to specific ranges in the "Queue Group" sheet to calculate HRM data.

2. **Pick and Replenishment Data**:

   - Aggregates data for picking and replenishment tasks and updates the "Data" sheet.

3. **Shift-Specific Data**:

   - Calculates and updates data for shifts A1, A2, and A3.

4. **Packing Data**:

   - Aggregates packing data and updates the "Data" sheet.

5. **Inbound Data**:

   - Aggregates inbound data and updates the "Data" sheet.

6. **Summary**:

   - Calculates summary metrics and updates the "Data" sheet.

7. **Highlight Specific Data**:
   - Highlights specific cells in the "Data" sheet based on conditions.

---

### Inputs

- **Data from Worksheets**:
  - The subroutine uses data from the "HRM", "P&R Lines", and "Data" sheets.

---

### Outputs

- **Updated Workbook**:
  - Updates the "Data" sheet with aggregated and calculated results.

---

### Limitations

1. **Hardcoded Ranges**:
   - The subroutine uses hardcoded ranges, which may need to be updated if the workbook structure changes.
2. **No Error Handling**:
   - The subroutine does not handle errors, such as missing data or invalid ranges.

---

## Key Components

### Subroutine: `Run`

| **Step**                | **Description**                                                        |
| ----------------------- | ---------------------------------------------------------------------- |
| Initialize Variables    | Resets all public variables to `0`.                                    |
| Clear Previous Data     | Clears specific ranges in the "Data" sheet and removes filters.        |
| Determine Current Day   | Uses the current date to determine the day of the week.                |
| Call Day-Specific Logic | Calls subroutines for Monday, Tuesday, etc., based on the current day. |

---

### Subroutine: `Consolidation`

| **Step**                | **Description**                                                    |
| ----------------------- | ------------------------------------------------------------------ |
| HRM Data                | Applies formulas to calculate HRM data.                            |
| Pick and Replenishment  | Aggregates data for picking and replenishment tasks.               |
| Shift-Specific Data     | Calculates and updates data for shifts A1, A2, and A3.             |
| Packing Data            | Aggregates packing data and updates the "Data" sheet.              |
| Inbound Data            | Aggregates inbound data and updates the "Data" sheet.              |
| Summary                 | Calculates summary metrics and updates the "Data" sheet.           |
| Highlight Specific Data | Highlights specific cells in the "Data" sheet based on conditions. |

---

## Assumptions

1. The workbook contains the required sheets ("HRM", "P&R Lines", "Data", etc.).
2. The data in the workbook is structured as expected.
3. The day-specific subroutines (`Monday`, `Tuesday`, etc.) are defined elsewhere in the project.

---

## Limitations

1. **Hardcoded Ranges**:

   - The subroutines rely on hardcoded ranges, which may need to be updated if the workbook structure changes.

2. **No Error Handling**:

   - The subroutines do not handle errors, such as missing sheets or invalid data.

3. **Dependency on External Subroutines**:
   - The `Run` subroutine depends on day-specific subroutines (`Monday`, `Tuesday`, etc.) that must be defined elsewhere.

---

## Recommendations for Improvement

1. **Add Error Handling**:

   - Include error handling to manage issues such as missing sheets or invalid data.  
     Example:
     ```vba
     On Error GoTo ErrorHandler
     ' Code here
     Exit Sub
     ErrorHandler:
     MsgBox "An error occurred: " & Err.Description
     ```

2. **Dynamic Ranges**:

   - Replace hardcoded ranges with dynamic range detection to make the subroutines more robust.

3. **Optimize Code**:
   - Remove unnecessary selections and streamline the code for better performance.

---

## Example Usage

### Run

1. Open the workbook containing the required sheets.
2. Run the `Run` subroutine.
3. Verify that the "Data" sheet is updated based on the current day of the week.

### Consolidation

1. Ensure that the "HRM", "P&R Lines", and "Data" sheets contain valid data.
2. Run the `Consolidation` subroutine.
3. Verify that the "Data" sheet is updated with aggregated and calculated results.

---
