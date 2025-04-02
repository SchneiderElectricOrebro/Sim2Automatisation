# SIM2 Macros Repository

This repository contains all the macros related to the SIM2 files. These macros are designed to automate and simplify various tasks associated with the SIM2 processes.

## Contents

- A collection of macros tailored for SIM2 workflows.
- Documentation and usage instructions for each macro.

## Purpose

The purpose of this repository is to centralize and manage all macros used for SIM2, ensuring consistency, efficiency, and ease of access for users.

## How to Use

1. Clone the repository to your local machine.
2. Follow the instructions provided in the documentation for each macro.
3. Execute the macros as needed to streamline your SIM2 tasks.

---

## Explanations of Each File

Here is the explanation of each file. To not lose time here is the explanation of the main used command :

-
-
-
-

### AccessExport

Purpose: The code exports data from specific cells in an Excel worksheet to a table named "SimSummary" in an Access database located at G:\09 Metod\14. Daily SIM Database\Daily SIM.accdb.

Steps:

-Connection Setup: It establishes a connection to the Access database using an ADODB.Connection object.
-Recordset Setup: It opens the "SimSummary" table in the database using an ADODB.Recordset object with dynamic cursor and optimistic locking.
-Data Transfer: It adds a new record to the "SimSummary" table and assigns values from specific Excel cells to the fields in this new record.
-Save and Close: It saves the new record to the database and closes the recordset.
-Notification: It displays a message box to inform the user that the data has been successfully exported.

Required Data:
The code needs data from the following Excel cell ranges:

W30, S13, S10, S11, S12, V30, E11, I11, M11, B19, B20, E10, I10, M10, B17, B18, B9
These cells should contain the data you want to export to the Access database. The data can be of various types (e.g., text, numbers) depending on what is stored in these cells.

---

### Import

ImportFiles Subroutine:

- Prompts the user to select an Excel file containing picking and replenish information.
- Opens the selected file and copies its visible sheets to the active workbook.
- Deletes any existing "P&R Lines" sheet in the active workbook and renames the copied sheet to "P&R Lines".
- Closes the opened workbook.

ImportHRM Subroutine:

- Prompts the user to select a text file containing HRM data.
- Creates a new sheet named "HRM" in the active workbook.
- Imports the data from the selected text file into the "HRM" sheet.
- Sets the first row of the "HRM" sheet to "N".
- Selects the "Data" sheet.

Required Data:

- ImportFiles: Requires an Excel file with picking and replenish information.
- ImportHRM: Requires a text file with HRM data.

---

### Module1

Macro1 Subroutine:

- Selects cell C2.
- Sets a formula in cell C2. The formula references cells in the "Queue Group" sheet and includes some invalid references (#REF!).

Macro2 Subroutine:

- Selects cell C2.
- Sets a formula in cell C2 that references cells in the "P&R Lines" and "Queue Group" sheets.
- Selects cell D2 and sets a similar formula.
- Selects cell E2 and sets a similar formula.
- Selects cell E3.

Required Data:

- Macro1: Needs data from the "Queue Group" sheet.
- Macro2: Needs data from the "P&R Lines" and "Queue Group" sheets.

---

### Module2

SD Subroutine:

- Displays a user form named SelectDate. This user form must be defined elsewhere in your VBA project.

test Subroutine:

- This subroutine is empty and does not perform any actions.

---

### Operator

#### Subroutine Name: Individual_Effectivity

##### Purpose:

The `Individual_Effectivity` subroutine calculates the individual effectivity of operators for a given period of time. It processes data from three worksheets (`Individual Performance`, `P&R Lines`, and `HRM`) to compute task counts, hours worked, and productivity metrics. The subroutine also calculates these metrics on a weekly basis.

##### Inputs:

1. Starting Week (`Wi`): The user is prompted to input the starting week for the calculation.
2. Ending Week (`Wf`): The user is prompted to input the ending week for the calculation.

##### Worksheets Used:

1. `Individual Performance` (InP):
   - Stores the results of the calculations, including task counts, hours worked, and productivity metrics.
2. `P&R Lines` (PRL):
   - Contains data about tasks performed by operators.
3. `HRM`:
   - Contains data about hours worked by operators.

##### Key Variables:

- `Wi` and `Wf`: Starting and ending weeks for the calculation.
- `arr`: An array that stores the SESAs (unique identifiers for operators) from the `Individual Performance` sheet.
- Task Counters:
  - `iOrdTruck`, `iHighLift`, `iSmalGang`, `iLongGoods`, `iPaternost`, `iRepl`: Counters for different types of tasks.
- HRM Data Variables:
  - `iHRMOrdTruck`, `iHRMHighLift`, `iHRMElevator`, `iHRMSmalgang`, `iHRMLongGoods`, `iHRMRepl`, `iHRMOthers`: Variables to store hours worked for different task types.

##### Steps:

1. Initialize Worksheets and Variables:

   - References to the `Individual Performance`, `P&R Lines`, and `HRM` worksheets are set.
   - The last row of data is determined for each worksheet.
   - The user is prompted to input the starting and ending weeks.

2. Extract SESAs:

   - SESAs (operator identifiers) are extracted from the `Individual Performance` sheet and stored in the `arr` array.

3. Clear Previous Data:

   - Clears the contents and formatting of the result columns in the `Individual Performance` sheet.

4. Process Data for the Total Period:

   - Loops through each operator in the arr array.
   - For each operator:
     - Resets task counters and HRM data variables.
     - Processes data from the `P&R Lines` sheet to count tasks performed by the operator.
     - Processes data from the `HRM` sheet to calculate hours worked by the operator.
     - Registers task counts, hours worked, and productivity metrics in the `Individual Performance` sheet.

5. Highlight Missing HRM Data:

   - Checks for missing HRM data and highlights cells with missing information in the `Individual Performance` sheet.

6. Weekly Calculations:

   - Loops through each week in the specified range (Wi to Wf).
   - For each week:
     - Resets task counters and HRM data variables.
     - Processes data from the `P&R Lines` and `HRM` sheets for the specific week.
     - Registers weekly task counts, hours worked, and productivity metrics in the `Individual Performance` sheet.

7. Highlight Missing Weekly HRM Data:

   - Checks for missing HRM data on a weekly basis and highlights cells with missing information.

##### Outputs:

The results are stored in the `Individual Performance` sheet:

1. Task Counts:

   - Total tasks performed by the operator.
   - Breakdown of tasks by type (e.g., ORD.TRUCK, HIGH LIFT, etc.).

2. Hours Worked:

   - Total hours worked by the operator.
   - Breakdown of hours by task type.

3. Productivity Metrics:

   - Productivity for each task type (tasks per hour).
   - Overall productivity.

##### Error Handling:

- The subroutine uses On Error Resume Next to handle potential errors during data processing. This ensures that the subroutine continues execution even if an error occurs.

##### Assumptions:

1. The `Individual Performance`, `P&R Lines`, and `HRM` worksheets exist and contain valid data.
2. The user inputs valid week numbers for the starting and ending weeks.
3. The data in the worksheets is structured as expected (e.g., specific columns contain specific types of data).

##### Limitations:

1. The subroutine does not validate the user input for the starting and ending weeks.
2. If the data structure in the worksheets changes, the subroutine may not function correctly.
3. The use of 'On Error Resume Next' may suppress critical errors, making debugging difficult.

---

## Contributions

Contributions are welcome! If you have improvements or new macros to add, please submit a pull request or open an issue.

## License

This repository is licensed under [insert license type]. Please review the license file for more details.
