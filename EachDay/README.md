# EachDay VBA Module Documentation

## Overview

The `EachDay.vba` file contains subroutines that process and analyze data for specific days of the week. Each subroutine (`Monday`, `Tuesday`, `Wednesday`, etc.) performs operations such as filtering, calculating shifts, aggregating task data, and preparing the data for consolidation. The file also includes a `WholeWeek` subroutine for processing data for the entire week.

---

## Subroutines

### 1. **Monday**

#### Purpose

Processes data for Monday by:

- Adding columns for exact time, weekday, and shift.
- Filtering rows based on specific conditions.
- Aggregating data for tasks such as picking, replenishment, and inbound.

#### Key Steps

1. **Add Columns**:

   - Adds columns for "Exact Hour confirmed", "Week day", and "Shift".
   - Determines the shift (`A1`, `A2`, `A3`) based on time and week number.

2. **Filter Rows**:

   - Deletes rows that do not match specific time and weekday conditions.

3. **Aggregate Data**:

   - Counts tasks such as "Ordinary Truck", "High Lift", "Small Gang", etc., and categorizes them by shift.

4. **Check Hours**:

   - Calculates hours for packing and inbound tasks.

5. **Call Consolidation**:
   - Calls the `Consolidation` subroutine to finalize data processing.

---

### 2. **Tuesday**

#### Purpose

Processes data for Tuesday using similar logic as the `Monday` subroutine but with Tuesday-specific conditions.

#### Key Steps

- Adds columns for time, weekday, and shift.
- Filters rows based on Tuesday-specific conditions.
- Aggregates task data and calculates hours for packing and inbound tasks.
- Calls the `Consolidation` subroutine.

---

### 3. **Wednesday**

#### Purpose

Processes data for Wednesday with Wednesday-specific conditions.

#### Key Steps

- Adds columns for time, weekday, and shift.
- Filters rows based on Wednesday-specific conditions.
- Aggregates task data and calculates hours for packing and inbound tasks.
- Calls the `Consolidation` subroutine.

---

### 4. **Thursday**

#### Purpose

Processes data for Thursday with Thursday-specific conditions.

#### Key Steps

- Adds columns for time, weekday, and shift.
- Filters rows based on Thursday-specific conditions.
- Aggregates task data and calculates hours for packing and inbound tasks.
- Calls the `Consolidation` subroutine.

---

### 5. **Friday**

#### Purpose

Processes data for Friday with Friday-specific conditions.

#### Key Steps

- Adds columns for time, weekday, and shift.
- Filters rows based on Friday-specific conditions.
- Aggregates task data and calculates hours for packing and inbound tasks.
- Calls the `Consolidation` subroutine.

---

### 6. **Weekend**

#### Purpose

Processes data for the weekend (Saturday and Sunday).

#### Key Steps

- Adds columns for time, weekday, and shift.
- Filters rows based on weekend-specific conditions.
- Aggregates task data and calculates hours for packing and inbound tasks.
- Calls the `Consolidation` subroutine.

---

### 7. **WholeWeek**

#### Purpose

Processes data for the entire week.

#### Key Steps

1. **Add Columns**:

   - Adds columns for "Exact Hour confirmed", "Week day", "Shift", and "Week Number".

2. **Filter Rows**:

   - Filters rows based on conditions for the entire week.

3. **Aggregate Data**:

   - Counts tasks such as "Ordinary Truck", "High Lift", "Small Gang", etc., and categorizes them by shift.

4. **Check Hours**:

   - Calculates hours for packing and inbound tasks.

5. **Call Consolidation**:
   - Calls the `Consolidation` subroutine to finalize data processing.

---

## Key Components

### Columns Added

| **Column Name**           | **Description**                                                        |
| ------------------------- | ---------------------------------------------------------------------- |
| `Exact Hour confirmed`    | Extracts the hour and minute from a timestamp.                         |
| `Week day`                | Determines the day of the week (1 = Sunday, 2 = Monday, etc.).         |
| `Shift`                   | Determines the shift (`A1`, `A2`, `A3`) based on time and week number. |
| `Week Number` (WholeWeek) | Determines the week number of the year.                                |

---

### Tasks Aggregated

| **Task Name**    | **Description**                                        |
| ---------------- | ------------------------------------------------------ |
| `Ordinary Truck` | Counts tasks related to ordinary trucks.               |
| `High Lift`      | Counts tasks related to high lifts.                    |
| `Small Gang`     | Counts tasks related to small gang operations.         |
| `Long Goods`     | Counts tasks related to long goods.                    |
| `Replenishment`  | Counts replenishment tasks (`REPL-HIGH`, `REPL-LONG`). |
| `Inbound`        | Counts inbound tasks based on specific conditions.     |

---

## Inputs

- **Sheets**:

  - `P&R Lines`: Contains task data for processing.
  - `HRM`: Contains hourly data for packing and inbound tasks.

- **Columns Used**:
  - `Column 19`: Timestamp for determining shifts.
  - `Column 18`: Date for determining weekdays and week numbers.
  - `Column 21`: Task type (e.g., "ORD.TRUCK", "HIGH LIFT").
  - `Column 6`: Task quantity.
  - `Column 15`: Task status.
  - `Column 22`: Task category.

---

## Outputs

- **Updated Sheets**:
  - Adds new columns to the `P&R Lines` and `HRM` sheets.
  - Updates task counts and hours in the `Data` sheet (via the `Consolidation` subroutine).

---

## Limitations

1. **Hardcoded Conditions**:

   - The subroutines rely on hardcoded conditions for filtering and aggregating data.

2. **No Error Handling**:

   - The subroutines do not handle errors, such as missing sheets or invalid data.

3. **Performance**:
   - The use of `Cells` and `Rows` for iterating through data may impact performance for large datasets.

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

### Daily Processing

1. Open the workbook containing the `P&R Lines` and `HRM` sheets.
2. Run the subroutine for the current day (e.g., `Monday`).
3. Verify that the data is processed and updated in the `Data` sheet.

### Weekly Processing

1. Open the workbook containing the `P&R Lines` and `HRM` sheets.
2. Run the `WholeWeek` subroutine.
3. Verify that the data for the entire week is processed and updated in the `Data` sheet.

---
