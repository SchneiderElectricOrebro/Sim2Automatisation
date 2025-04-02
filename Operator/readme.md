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
