# AccessExport Subroutine Documentation

## Purpose

The `AccessExport` subroutine exports data from specific Excel ranges into an Access database table (`SimSummary`). It is designed to save daily data into the database for record-keeping and analysis.

---

## Inputs

### Excel Ranges

- The subroutine reads data from specific cells in the active Excel worksheet (e.g., `W30`, `S13`, `S10`, etc.).
- These ranges are mapped to fields in the `SimSummary` table in the Access database.

### Access Database

- The database file is located at:  
  `G:\09 Metod\14. Daily SIM Database\Daily SIM.accdb`.

---

## Outputs

### Access Table (`SimSummary`)

- A new record is added to the `SimSummary` table in the Access database.
- The fields in the table are populated with the values from the specified Excel ranges.

### Message Box

- A message box is displayed to confirm that the data has been successfully exported.

---

## Steps

1. **Initialize ADODB Objects**:

   - A new `ADODB.Connection` object (`newcon`) is created to establish a connection to the Access database.
   - A new `ADODB.Recordset` object (`Recordset`) is created to interact with the `SimSummary` table.

2. **Open Database Connection**:

   - The `newcon.Open` method is used to connect to the Access database using the `Microsoft.ACE.OLEDB.12.0` provider.

3. **Open the `SimSummary` Table**:

   - The `Recordset.Open` method is used to open the `SimSummary` table with a dynamic cursor (`adOpenDynamic`) and optimistic locking (`adLockOptimistic`).

4. **Add a New Record**:

   - The `Recordset.AddNew` method is called to create a new record in the `SimSummary` table.

5. **Assign Values to Fields**:

   - Data from specific Excel ranges is assigned to the fields in the new record using the `Recordset.Fields` property.  
     Example:  
     `Recordset.Fields(1).Value = Range("W30").Value` assigns the value from cell `W30` to the first field in the table.

6. **Save the Record**:

   - The `Recordset.Update` method is called to save the new record to the database.

7. **Close the Recordset**:

   - The `Recordset.Close` method is called to close the recordset and release resources.

8. **Display Confirmation Message**:
   - A message box is displayed to inform the user that the data has been successfully exported.

---

## Key Components

### Excel Ranges Used

| **Field Number** | **Excel Range** | **Description**                              |
| ---------------- | --------------- | -------------------------------------------- |
| 1                | `W30`           | Data for the first field in the table.       |
| 2                | `S13`           | Data for the second field in the table.      |
| 3                | `S10`           | Data for the third field in the table.       |
| 4                | `S11`           | Data for the fourth field in the table.      |
| 5                | `S12`           | Data for the fifth field in the table.       |
| 6                | `V30`           | Data for the sixth field in the table.       |
| 7                | `E11`           | Data for the seventh field in the table.     |
| 8                | `I11`           | Data for the eighth field in the table.      |
| 9                | `M11`           | Data for the ninth field in the table.       |
| 10               | `B19`           | Data for the tenth field in the table.       |
| 11               | `B20`           | Data for the eleventh field in the table.    |
| 12               | `E10`           | Data for the twelfth field in the table.     |
| 13               | `I10`           | Data for the thirteenth field in the table.  |
| 14               | `M10`           | Data for the fourteenth field in the table.  |
| 15               | `B17`           | Data for the fifteenth field in the table.   |
| 16               | `B18`           | Data for the sixteenth field in the table.   |
| 17               | `B9`            | Data for the seventeenth field in the table. |

---

## Error Handling

- The subroutine does not include explicit error handling. If an error occurs (e.g., the database file is missing or the Excel ranges are invalid), the subroutine will terminate with a runtime error.

---

## Assumptions

1. The Access database file (`Daily SIM.accdb`) exists at the specified path.
2. The `SimSummary` table exists in the database and has at least 17 fields.
3. The Excel ranges specified in the subroutine contain valid data.
4. The `Microsoft.ACE.OLEDB.12.0` provider is installed and available on the system.

---

## Limitations

1. **Hardcoded File Path**:

   - The database file path is hardcoded (`G:\09 Metod\14. Daily SIM Database\Daily SIM.accdb`). If the file is moved or renamed, the subroutine will fail.

2. **No Error Handling**:

   - The subroutine does not handle errors, such as missing files, invalid ranges, or database connection issues.

3. **Field Mapping**:
   - The mapping between Excel ranges and database fields is fixed. Changes to the database schema or Excel layout will require updates to the subroutine.

---

## Example Usage

1. Open the Excel workbook containing the subroutine.
2. Ensure that the required data is present in the specified ranges (e.g., `W30`, `S13`, etc.).
3. Run the `AccessExport` subroutine.
4. Verify that the data has been added to the `SimSummary` table in the Access database.

---

## Recommendations for Improvement

1. **Error Handling**:

   - Add error handling to manage issues such as missing files, invalid ranges, or database connection errors.  
     Example:
     ```vba
     On Error GoTo ErrorHandler
     ' Code here
     Exit Sub
     ErrorHandler:
     MsgBox "An error occurred: " & Err.Description
     ```

2. **Dynamic File Path**:

   - Allow the user to specify the database file path dynamically instead of hardcoding it.

3. **Validation**:
   - Validate the data in the Excel ranges before exporting it to the database.

---
