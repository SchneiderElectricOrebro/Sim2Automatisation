Sub AccessExport()

    ' This subroutine saves the main data into an Access table according to the day

    ' Declare a new ADODB.Connection object
    Dim newcon As ADODB.Connection
    Set newcon = New ADODB.Connection

    ' Declare a new ADODB.Recordset object
    Dim Recordset As ADODB.Recordset
    Set Recordset = New ADODB.Recordset

    ' Open a connection to the Access database
    newcon.Open "provider=microsoft.ace.oledb.12.0;Data source= G:\09 Metod\14. Daily SIM Database\Daily SIM.accdb"

    ' Open the "SimSummary" table in the database with dynamic cursor and optimistic locking
    Recordset.Open "SimSummary", newcon, adOpenDynamic, adLockOptimistic

    ' Add a new record to the recordset
    Recordset.AddNew

    ' Assign values from specific Excel ranges to the fields in the new record
    Recordset.Fields(1).Value = Range("W30").Value
    Recordset.Fields(2).Value = Range("S13").Value
    Recordset.Fields(3).Value = Range("S10").Value
    Recordset.Fields(4).Value = Range("S11").Value
    Recordset.Fields(5).Value = Range("S12").Value
    Recordset.Fields(6).Value = Range("V30").Value
    Recordset.Fields(7).Value = Range("E11").Value
    Recordset.Fields(8).Value = Range("I11").Value
    Recordset.Fields(9).Value = Range("M11").Value
    Recordset.Fields(10).Value = Range("B19").Value
    Recordset.Fields(11).Value = Range("B20").Value
    Recordset.Fields(12).Value = Range("E10").Value
    Recordset.Fields(13).Value = Range("I10").Value
    Recordset.Fields(14).Value = Range("M10").Value
    Recordset.Fields(15).Value = Range("B17").Value
    Recordset.Fields(16).Value = Range("B18").Value
    Recordset.Fields(17).Value = Range("B9").Value

    ' Save the new record to the database
    Recordset.Update

    ' Close the recordset
    Recordset.Close

    ' Display a message box to inform the user that the data has been exported
    MsgBox "Data Exported to Sim Summary Table: G:\09 Metod\14. Daily SIM Database\Daily SIM.accdb"

End Sub