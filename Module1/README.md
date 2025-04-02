# Module1 VBA Macros Documentation

## Overview

This file contains two VBA macros, `Macro1` and `Macro2`, which are designed to set formulas in specific cells of an Excel worksheet. These macros automate the process of applying formulas to cells, but the provided formulas appear to contain errors (`#REF!`), which need to be corrected for proper functionality.

---

## Subroutine: `Macro1`

### Purpose

The `Macro1` subroutine sets a formula in cell `C2` of the active worksheet. However, the formula contains unresolved references (`#REF!`), which need to be fixed.

---

### Steps

1. **Select Cell `C2`**:

   - The macro selects cell `C2` in the active worksheet.

2. **Set Formula**:

   - A formula is applied to cell `C2` using the `FormulaR1C1` property.  
     **Note**: The formula contains unresolved references (`#REF!`), which need to be corrected.

3. **Select Cell `C2` Again**:
   - The macro redundantly selects cell `C2` again, which is unnecessary since it is already selected.

---

### Inputs

- **Active Worksheet**:
  - The macro operates on the currently active worksheet.

---

### Outputs

- **Cell `C2`**:
  - The formula is applied to cell `C2` of the active worksheet.

---

### Limitations

1. **Unresolved References**:
   - The formula contains `#REF!`, which indicates missing or invalid references.
2. **Redundant Selection**:
   - The macro selects cell `C2` twice, which is unnecessary.

---

## Subroutine: `Macro2`

### Purpose

The `Macro2` subroutine sets formulas in cells `C2`, `D2`, and `E2` of the active worksheet. Similar to `Macro1`, the formulas contain unresolved references (`#REF!`), which need to be corrected.

---

### Steps

1. **Select Cell `C2`**:

   - The macro selects cell `C2` in the active worksheet.

2. **Set Formula in `C2`**:

   - A formula is applied to cell `C2` using the `FormulaR1C1` property.  
     **Note**: The formula contains unresolved references (`#REF!`), which need to be corrected.

3. **Select Cell `D2`**:

   - The macro selects cell `D2` in the active worksheet.

4. **Set Formula in `D2`**:

   - A formula is applied to cell `D2` using the `FormulaR1C1` property.  
     **Note**: The formula contains unresolved references (`#REF!`), which need to be corrected.

5. **Select Cell `E2`**:

   - The macro selects cell `E2` in the active worksheet.

6. **Set Formula in `E2`**:

   - A formula is applied to cell `E2` using the `FormulaR1C1` property.  
     **Note**: The formula contains unresolved references (`#REF!`), which need to be corrected.

7. **Select Cell `E3`**:
   - The macro selects cell `E3` in the active worksheet.

---

### Inputs

- **Active Worksheet**:
  - The macro operates on the currently active worksheet.

---

### Outputs

- **Cells `C2`, `D2`, and `E2`**:
  - Formulas are applied to these cells in the active worksheet.

---

### Limitations

1. **Unresolved References**:
   - The formulas contain `#REF!`, which indicates missing or invalid references.
2. **Redundant Selection**:
   - The macro selects cells unnecessarily, which can be optimized.

---

## Key Components

### Subroutine: `Macro1`

| **Step**            | **Description**                                |
| ------------------- | ---------------------------------------------- |
| Select Cell `C2`    | Selects cell `C2` in the active worksheet.     |
| Set Formula in `C2` | Applies a formula to cell `C2`.                |
| Redundant Selection | Selects cell `C2` again, which is unnecessary. |

---

### Subroutine: `Macro2`

| **Step**            | **Description**                            |
| ------------------- | ------------------------------------------ |
| Select Cell `C2`    | Selects cell `C2` in the active worksheet. |
| Set Formula in `C2` | Applies a formula to cell `C2`.            |
| Select Cell `D2`    | Selects cell `D2` in the active worksheet. |
| Set Formula in `D2` | Applies a formula to cell `D2`.            |
| Select Cell `E2`    | Selects cell `E2` in the active worksheet. |
| Set Formula in `E2` | Applies a formula to cell `E2`.            |
| Select Cell `E3`    | Selects cell `E3` in the active worksheet. |

---

## Assumptions

1. The active worksheet contains valid data and references for the formulas.
2. The unresolved references (`#REF!`) in the formulas will be corrected before running the macros.

---

## Limitations

1. **Unresolved References**:
   - The formulas contain `#REF!`, which indicates missing or invalid references.
2. **Redundant Selections**:
   - The macros select cells unnecessarily, which can be optimized for better performance.
3. **No Error Handling**:
   - The macros do not handle errors, such as invalid references or missing data.

---

## Recommendations for Improvement

1. **Fix Formulas**:

   - Correct the unresolved references (`#REF!`) in the formulas to ensure proper functionality.

2. **Optimize Code**:

   - Remove redundant cell selections to improve performance and readability.

3. **Add Error Handling**:
   - Include error handling to manage issues such as invalid references or missing data.  
     Example:
     ```vba
     On Error GoTo ErrorHandler
     ' Code here
     Exit Sub
     ErrorHandler:
     MsgBox "An error occurred: " & Err.Description
     ```

---

## Example Usage

### Macro1

1. Run the `Macro1` subroutine.
2. Verify that the formula is applied to cell `C2` in the active worksheet.

### Macro2

1. Run the `Macro2` subroutine.
2. Verify that the formulas are applied to cells `C2`, `D2`, and `E2` in the active worksheet.

---
