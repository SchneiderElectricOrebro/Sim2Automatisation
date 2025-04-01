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

## Contributions

Contributions are welcome! If you have improvements or new macros to add, please submit a pull request or open an issue.

## License

This repository is licensed under [insert license type]. Please review the license file for more details.
