# Operators VBA Module Documentation

## Overview

The `Operators.vba` file contains subroutines for calculating individual operator performance and productivity metrics. The main subroutine, `Individual_Effectivity`, processes data from multiple worksheets to calculate task counts, hours worked, and productivity for each operator over a specified period. It also includes weekly calculations and highlights missing HRM data.

---

## Subroutines

### 1. **Individual_Effectivity**

#### Purpose

Calculates individual operator performance and productivity for a specified period and on a weekly basis. It processes data from the `P&R Lines` and `HRM` sheets and stores the results in the `Individual Performance` sheet.

---

#### Key Steps

1. **Initialize Variables**:

   - Declares and initializes variables for task counts, HRM data, and operator identifiers (SESAs).

2. **Prompt for Weeks**:

   - Prompts the user to input the starting (`Wi`) and ending (`Wf`) weeks for the calculation.

3. **Process Data from `P&R Lines`**:

   - Loops through the `P&R Lines` sheet to count tasks such as "Ordinary Truck", "High Lift", "Small Gang", etc., for each operator.

4. **Process Data from `HRM`**:

   - Loops through the `HRM` sheet to calculate hours worked for each task type.

5. **Register Results**:

   - Updates the `Individual Performance` sheet with task counts, hours worked, and productivity metrics for each operator.

6. **Highlight Missing HRM Data**:

   - Highlights cells where task data exists but corresponding HRM data is missing.

7. **Weekly Calculations**:
   - Repeats the above steps for each week in the specified range (`Wi` to `Wf`).

---

#### Inputs

- **User Input**:
  - Starting week (`Wi`) and ending week (`Wf`).
- **Worksheets**:
  - `Individual Performance`: Stores the results.
  - `P&R Lines`: Contains task data.
  - `HRM`: Contains hours worked data.

---

#### Outputs

- **Updated `Individual Performance` Sheet**:
  - Task counts, hours worked, and productivity metrics for each operator.
  - Weekly data for the specified period.

---

#### Limitations

1. **Hardcoded Conditions**:
   - The subroutine uses hardcoded conditions for filtering and aggregating data.
2. **No Error Handling**:
   - The subroutine does not handle errors, such as missing sheets or invalid data.

---

### 2. **Individual_EffectivityTotal**

#### Purpose

Calculates individual operator performance and productivity for the total period of extraction. It processes data from the `P&R Lines` and `HRM` sheets and stores the results in the `Individual Performance` sheet.

---

#### Key Steps

1. **Initialize Variables**:

   - Declares and initializes variables for task counts, HRM data, and operator identifiers (SESAs).

2. **Process Data from `P&R Lines`**:

   - Loops through the `P&R Lines` sheet to count tasks such as "Ordinary Truck", "High Lift", "Small Gang", etc., for each operator.

3. **Process Data from `HRM`**:

   - Loops through the `HRM` sheet to calculate hours worked for each task type.

4. **Register Results**:

   - Updates the `Individual Performance` sheet with task counts, hours worked, and productivity metrics for each operator.

5. **Highlight Missing HRM Data**:
   - Highlights cells where task data exists but corresponding HRM data is missing.

---

#### Inputs

- **Worksheets**:
  - `Individual Performance`: Stores the results.
  - `P&R Lines`: Contains task data.
  - `HRM`: Contains hours worked data.

---

#### Outputs

- **Updated `Individual Performance` Sheet**:
  - Task counts, hours worked, and productivity metrics for each operator.

---

#### Limitations

1. **Hardcoded Conditions**:
   - The subroutine uses hardcoded conditions for filtering and aggregating data.
2. **No Error Handling**:
   - The subroutine does not handle errors, such as missing sheets or invalid data.

---

## Key Components

### Worksheets Used

| **Worksheet**            | **Description**                                                                                        |
| ------------------------ | ------------------------------------------------------------------------------------------------------ |
| `Individual Performance` | Stores the results of the calculations, including task counts, hours worked, and productivity metrics. |
| `P&R Lines`              | Contains task data for operators.                                                                      |
| `HRM`                    | Contains hours worked data for operators.                                                              |

---

### Tasks Aggregated

| **Task Name**    | **Description**                                        |
| ---------------- | ------------------------------------------------------ |
| `Ordinary Truck` | Counts tasks related to ordinary trucks.               |
| `High Lift`      | Counts tasks related to high lifts.                    |
| `Small Gang`     | Counts tasks related to small gang operations.         |
| `Long Goods`     | Counts tasks related to long goods.                    |
| `Paternoster`    | Counts tasks related to paternoster operations.        |
| `Replenishment`  | Counts replenishment tasks (`REPL-HIGH`, `REPL-LONG`). |

---

### HRM Data Processed

| **HRM Task Type** | **Description**                       |
| ----------------- | ------------------------------------- |
| `Ordinary Truck`  | Hours worked on ordinary truck tasks. |
| `High Lift`       | Hours worked on high lift tasks.      |
| `Elevator`        | Hours worked on elevator tasks.       |
| `Small Gang`      | Hours worked on small gang tasks.     |
| `Long Goods`      | Hours worked on long goods tasks.     |
| `Replenishment`   | Hours worked on replenishment tasks.  |
| `Others`          | Hours worked on other tasks.          |

---

## Recommendations for Improvement

1. **Add Error Handling**:

   - Include error handling to manage issues such as missing sheets or invalid data.  
     Example:
     ```vba
     On Error Resume Next
     ' Code here
     If Err.Number <> 0 Then
         MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
     End If
     On Error GoTo 0
     ```

2. **Optimize Code**:

   - Use `Range` objects and arrays for faster data processing instead of iterating through rows with `Cells`.

3. **Dynamic Conditions**:
   - Replace hardcoded conditions with configurable parameters to make the subroutines more flexible.

---

## Example Usage

### Individual_Effectivity

1. Open the workbook containing the `Individual Performance`, `P&R Lines`, and `HRM` sheets.
2. Run the `Individual_Effectivity` subroutine.
3. Input the starting and ending weeks when prompted.
4. Verify that the `Individual Performance` sheet is updated with task counts, hours worked, and productivity metrics.

### Individual_EffectivityTotal

1. Open the workbook containing the `Individual Performance`, `P&R Lines`, and `HRM` sheets.
2. Run the `Individual_EffectivityTotal` subroutine.
3. Verify that the `Individual Performance` sheet is updated with task counts, hours worked, and productivity metrics for the total period.

---

Let me know if you need further modifications or additional details!
