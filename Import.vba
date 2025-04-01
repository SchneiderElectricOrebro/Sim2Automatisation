Public Sub ImportFiles()

' Find the extractions from SAP and copy their location
Dim wb1 As Workbook
Dim wbPickedLines As Workbook

Set wb1 = ActiveWorkbook

'FileToOpen is the excel file containing picking and replenish information
FileToOpen = Application.GetOpenFilename _
(Title:="Please choose the Picked Lines extraction", _
FileFilter:="Report Files *.xls (*.xlsx, .xls),")

If FileToOpen = False Then
    MsgBox "No file selected.", vbExclamation, "Erro"
    Exit Sub
Else
    Set wbPickedLines = Workbooks.Open(Filename:=FileToOpen)


    For Each Sheet In wbPickedLines.Sheets
        If Sheet.Visible = True Then
            On Error Resume Next
            Application.DisplayAlerts = False
            Sheet.Copy After:=wb1.Sheets(1)
            wb1.Worksheets("P&R Lines").Delete
            ActiveSheet.Name = "P&R Lines"
        End If
    Next Sheet

End If

    wbPickedLines.Close
    
Dim wPL As Worksheet: Set wPL = Sheets("P&R Lines")

wb1.Select
Worksheets("Data").Select
    
End Sub
    
Public Sub ImportHRM()


'TxTToOpen is the text containing HRM Data
TxTToOpen = Application.GetOpenFilename _
(Title:="Please choose the HRM extraction", _
FileFilter:="Report Files *.txt (*.txt),")

If TxTToOpen <> False Then
        
        'Create sheet
        On Error Resume Next
        Application.DisplayAlerts = False
        Worksheets("HRM").Delete
        Sheets.Add After:=Sheets("Data")
        ActiveSheet.Name = "HRM"
        Application.DisplayAlerts = True
        Set wHRM = Worksheets("HRM")
        
        'Import TXT
        
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
        
        'Transformar colunas pra texto
        'Columns("H:H").Select
    'Selection.TextToColumns Destination:=Range("H2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
        
    Else
        MsgBox "No file selected.", vbExclamation, "Erro"
        Exit Sub
End If


wHRM.Range("A1:J1").Value = "N"
Worksheets("Data").Select
    
Application.DisplayAlerts = False


End Sub


