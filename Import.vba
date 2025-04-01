Public Sub ImportFiles()

    ' Find the extractions from SAP and copy their location

    ' Declare workbook variables
    Dim wb1 As Workbook
    Dim wbPickedLines As Workbook

    ' Set the active workbook to wb1
    Set wb1 = ActiveWorkbook

    ' Prompt the user to select the Excel file containing picking and replenish information
    FileToOpen = Application.GetOpenFilename( _
        Title:="Please choose the Picked Lines extraction", _
        FileFilter:="Report Files *.xls (*.xlsx, .xls),")

    ' Check if a file was selected
    If FileToOpen = False Then
        MsgBox "No file selected.", vbExclamation, "Error"
        Exit Sub
    Else
        ' Open the selected workbook
        Set wbPickedLines = Workbooks.Open(Filename:=FileToOpen)

        ' Loop through each sheet in the opened workbook
        For Each Sheet In wbPickedLines.Sheets
            If Sheet.Visible = True Then
                On Error Resume Next
                Application.DisplayAlerts = False
                ' Copy the sheet to the active workbook
                Sheet.Copy After:=wb1.Sheets(1)
                ' Delete the existing "P&R Lines" sheet in the active workbook
                wb1.Worksheets("P&R Lines").Delete
                ' Rename the copied sheet to "P&R Lines"
                ActiveSheet.Name = "P&R Lines"
            End If
        Next Sheet

    End If

    ' Close the opened workbook
    wbPickedLines.Close

    ' Set the "P&R Lines" sheet to a variable
    Dim wPL As Worksheet: Set wPL = Sheets("P&R Lines")

    ' Select the active workbook and the "Data" sheet
    wb1.Select
    Worksheets("Data").Select

End Sub

Public Sub ImportHRM()

    ' Prompt the user to select the text file containing HRM data
    TxTToOpen = Application.GetOpenFilename( _
        Title:="Please choose the HRM extraction", _
        FileFilter:="Report Files *.txt (*.txt),")

    ' Check if a file was selected
    If TxTToOpen <> False Then

        ' Create a new sheet for HRM data
        On Error Resume Next
        Application.DisplayAlerts = False
        Worksheets("HRM").Delete
        Sheets.Add After:=Sheets("Data")
        ActiveSheet.Name = "HRM"
        Application.DisplayAlerts = True
        Set wHRM = Worksheets("HRM")

        ' Import the text file data into the new sheet
        With ActiveSheet.QueryTables.Add(Connection:= _
            "TEXT;" & TxTToOpen, Destination:=Range("$A$2"))
            .CommandType = 0
            .Name = "HRM Report"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .TextFileSemicolonDelimiter = True
            .Refresh BackgroundQuery:=False
        End With

    Else
        MsgBox "No file selected.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Set the first row of the HRM sheet to "N"
    wHRM.Range("A1:J1").Value = "N"

    ' Select the "Data" sheet
    Worksheets("Data").Select

    Application.DisplayAlerts = False

End Sub