Sub Individual_Effectivity()
'This calculates the individual effectivity for the total period of extraction and also calculates on a weekly basis

Dim InP As Worksheet
Set InP = Worksheets("Individual Performance")

Dim PRL As Worksheet
Set PRL = Worksheets("P&R Lines")

Dim HRM As Worksheet
Set HRM = Worksheets("HRM")

Dim LastRowInP As Range
Set LastRowInP = InP.Range("A900000").End(xlUp)

Dim W As Long

Dim Wi As Long
Dim Wf As Long

Wi = InputBox("Please enter the starting week")
Wf = InputBox("Please enter the final week")



'Declaring SESAs as strings in the Individual sheet
Dim arr(1 To 127) As String

Dim K As Integer
For K = 3 To LastRowInP.Row
    arr(K - 2) = InP.Cells(K, 1).Value
Next



Dim i As Long
Set LastRowPL = PRL.Range("A900000").End(xlUp)
Set LastRowHRM = HRM.Range("A900000").End(xlUp)

'Dim of each variable
Dim iOrdTruck As Long
Dim iHighLift As Long
Dim iSmalGang As Long
Dim iLongGoods As Long
Dim iPaternost As Long
Dim iRepl As Long
Dim iTest As Long

Dim iR As Long

Dim iHRMOrdTruck As Double
Dim iHRMHighLift As Double
Dim iHRMElevator As Double
Dim iHRMSmalgang As Double
Dim iHRMLongGoods As Double
Dim iHRMRepl As Double



'Clean Sheet
InP.Range("F3:JI" & LastRowInP.Row).ClearContents
InP.Range("F3:JI" & LastRowInP.Row).Interior.Color = xlNone

'**********************************************************************************************************
'Get the total measured in the first set of data


'Loop items in array
For i = LBound(arr) To UBound(arr)

'Gathering of data in P&L Rep sheet------------------------------------------------------------
'Reset of variables


iOrdTruck = 0
iHighLift = 0
iSmalGang = 0
iLongGoods = 0
iPaternost = 0
iRepl = 0

iHRMOrdTruck = 0
iHRMHighLift = 0
iHRMElevator = 0
iHRMSmalgang = 0
iHRMElevator = 0
iHRMLongGoods = 0
iHRMRepl = 0
iHRMOthers = 0




    'PRL.Activate
    For iRow = 3 To LastRowPL.Row
    
    
        'Check if the line works for appointed operator
        If arr(i) = PRL.Cells(iRow, 17).Value Then

        
        
            If (PRL.Cells(iRow, 15).Value = 100 Or PRL.Cells(iRow, 15).Value = 916) And (PRL.Cells(iRow, 22).Value <> 20 Or PRL.Cells(iRow, 22).Value <> 21 Or PRL.Cells(iRow, 22).Value <> 120 Or PRL.Cells(iRow, 22).Value <> 121) Then
                'Count the total of picked Lines
                If (PRL.Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(PRL.Cells(iRow, 21).Value, 3) = "DPI" Or Left(PRL.Cells(iRow, 21).Value, 3) = "FBO" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAF" Then
                    iOrdTruck = iOrdTruck + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "ORD.ELKO") Or Left(PRL.Cells(iRow, 21).Value, 3) = "DPI" Or Left(PRL.Cells(iRow, 21).Value, 3) = "FBO" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAF" Then
                    iOrdTruck = iOrdTruck + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "HIGH LIFT") Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRP" Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRF" Then
                    iHighLift = iHighLift + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "SMALGANG 1") Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAF" Then
                    iSmalGang = iSmalGang + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "SMALGANG_E") Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAF" Then
                    iSmalGang = iSmalGang + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "LONG GOODS") Then
                    iLongGoods = iLongGoods + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "PATERNOST.") Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAT" Then
                    iPaternost = iPaternost + 1
                End If
            ElseIf PRL.Cells(iRow, 21).Value = "REPL-HIGH" Or PRL.Cells(iRow, 21).Value = "REPL-LONG" Then
                iRepl = iRepl + 1
                
            End If
        
       End If

        
    Next iRow

    'Registering picking data for specific operator:
    InP.Cells(i + 2, 6).Value = iOrdTruck + iHighLift + iSmalGang + iLongGoods + iPaternost
    InP.Cells(i + 2, 17).Value = iRepl
    InP.Cells(i + 2, 7).Value = iOrdTruck
    InP.Cells(i + 2, 9).Value = iHighLift
    InP.Cells(i + 2, 11).Value = iPaternost
    InP.Cells(i + 2, 13).Value = iSmalGang
    InP.Cells(i + 2, 15).Value = iLongGoods

     
    
'Getting data from HRM File-----------------------


    
    For iRow = 2 To LastRowHRM.Row
    

    
        If arr(i) = HRM.Cells(iRow, 2).Value Then
            
            If HRM.Cells(iRow, 3).Value = 600 Or HRM.Cells(iRow, 3).Value = 604 Or HRM.Cells(iRow, 3).Value = 608 Or HRM.Cells(iRow, 3).Value = 617 Or HRM.Cells(iRow, 3).Value = 629 Or HRM.Cells(iRow, 3).Value = 630 Then
                iHRMOrdTruck = iHRMOrdTruck + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 601 Or HRM.Cells(iRow, 3).Value = 605 Or HRM.Cells(iRow, 3).Value = 609 Then
                iHRMHighLift = iHRMHighLift + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 603 Or HRM.Cells(iRow, 3).Value = 607 Or HRM.Cells(iRow, 3).Value = 611 Then
                iHRMElevator = iHRMElevator + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 602 Or HRM.Cells(iRow, 3).Value = 606 Or HRM.Cells(iRow, 3).Value = 618 Then
                iHRMSmalgang = iHRMSmalgang + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 616 Then
                iHRMLongGoods = iHRMLongGoods + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 628 Or HRM.Cells(iRow, 3).Value = 653 Then
                iHRMRepl = iHRMRepl + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value <> "" And HRM.Cells(iRow, 5).Value <> "RAST" Then
                iHRMOthers = iHRMOthers + HRM.Cells(iRow, 11).Value
            End If
        End If
    Next
    
        

On Error Resume Next

'Registering Hours and productivity
    'Hours

    InP.Cells(i + 2, 18).Value = iHRMRepl
    InP.Cells(i + 2, 8).Value = iHRMOrdTruck
    InP.Cells(i + 2, 10).Value = iHRMHighLift
    InP.Cells(i + 2, 12).Value = iHRMElevator
    InP.Cells(i + 2, 14).Value = iHRMSmalgang
    InP.Cells(i + 2, 16).Value = iHRMLongGoods
    'Total Hours
    InP.Cells(i + 2, 19).Value = (iHRMOrdTruck + iHRMHighLift + iHRMElevator + iHRMSmalgang + iHRMLongGoods)
    InP.Cells(i + 2, 20).Value = iHRMRepl
    InP.Cells(i + 2, 21).Value = iHRMOthers
    
    'Productivity
    InP.Cells(i + 2, 22).Value = (iOrdTruck + iHighLift + iSmalGang + iLongGoods + iPaternost) / (iHRMOrdTruck + iHRMHighLift + iHRMElevator + iHRMSmalgang + iHRMLongGoods)
    InP.Cells(i + 2, 27).Value = iRepl / iHRMRepl
    InP.Cells(i + 2, 23).Value = iOrdTruck / iHRMOrdTruck
    InP.Cells(i + 2, 24).Value = iHighLift / iHRMHighLift
    InP.Cells(i + 2, 25).Value = iSmalGang / iHRMSmalgang
    InP.Cells(i + 2, 26).Value = iLongGoods / iHRMLongGoods
    
    
    If InP.Cells(i + 2, 12).Value > 0 Then InP.Cells(i + 2, 12).Interior.ColorIndex = 50
    

 
'Check cells  where there are missing HRM

For j = 8 To 18 Step 2
    For iR = 3 To LastRowInP.Row
    
    If InP.Cells(iR, j - 1).Value > 0 And InP.Cells(iR, j).Value = 0 Then
        InP.Cells(iR, j).Value = "No HRM Info"
        InP.Cells(iR, j).Interior.ColorIndex = 44
    End If
    Next
Next


Next i





'Now get it weekly
'************************************************************************************************************
'Loop from the LowerBound to UpperBound items in array
For i = LBound(arr) To UBound(arr)

'Gathering of data in P&L Rep sheet------------------------------------------------------------
'Reset of variables

'Weekly Code added
For W = Wi To Wf 'W


iOrdTruck = 0
iHighLift = 0
iSmalGang = 0
iLongGoods = 0
iPaternost = 0
iRepl = 0

iHRMOrdTruck = 0
iHRMHighLift = 0
iHRMElevator = 0
iHRMSmalgang = 0
iHRMElevator = 0
iHRMLongGoods = 0
iHRMRepl = 0
iHRMOthers = 0






    'PRL.Activate
    For iRow = 3 To LastRowPL.Row
    
      If PRL.Cells(iRow, 26).Value = W Then 'W
    
        'Check if the line works for appointed operator
        If arr(i) = PRL.Cells(iRow, 17).Value Then
        'iTest = iTest + 1
        'PRL.Cells(iRow, 26).Value = PRL.Cells(iRow, 17).Value
        
        
            If (PRL.Cells(iRow, 15).Value = 100 Or PRL.Cells(iRow, 15).Value = 916) And (PRL.Cells(iRow, 22).Value <> 20 Or PRL.Cells(iRow, 22).Value <> 21 Or PRL.Cells(iRow, 22).Value <> 120 Or PRL.Cells(iRow, 22).Value <> 121) Then
                'Count the total of picked Lines
                If (PRL.Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(PRL.Cells(iRow, 21).Value, 3) = "DPI" Or Left(PRL.Cells(iRow, 21).Value, 3) = "FBO" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAF" Then
                    iOrdTruck = iOrdTruck + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "ORD.ELKO") Or Left(PRL.Cells(iRow, 21).Value, 3) = "DPI" Or Left(PRL.Cells(iRow, 21).Value, 3) = "FBO" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAF" Then
                    iOrdTruck = iOrdTruck + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "HIGH LIFT") Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRP" Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRF" Then
                    iHighLift = iHighLift + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "SMALGANG 1") Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAF" Then
                    iSmalGang = iSmalGang + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "SMALGANG_E") Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAF" Then
                    iSmalGang = iSmalGang + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "LONG GOODS") Then
                    iLongGoods = iLongGoods + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "PATERNOST.") Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAT" Then
                    iPaternost = iPaternost + 1
                End If
            ElseIf PRL.Cells(iRow, 21).Value = "REPL-HIGH" Or PRL.Cells(iRow, 21).Value = "REPL-LONG" Then
                iRepl = iRepl + 1
                
            End If
        
       End If
     End If 'W
        
    Next iRow

    'Registering picking data for specific operator:
    InP.Cells(i + 2, 6 + (1 + W - Wi) * 22).Value = iOrdTruck + iHighLift + iSmalGang + iLongGoods + iPaternost
    InP.Cells(i + 2, 17 + (1 + W - Wi) * 22).Value = iRepl
    InP.Cells(i + 2, 7 + (1 + W - Wi) * 22).Value = iOrdTruck
    InP.Cells(i + 2, 9 + (1 + W - Wi) * 22).Value = iHighLift
    InP.Cells(i + 2, 11 + (1 + W - Wi) * 22).Value = iPaternost
    InP.Cells(i + 2, 13 + (1 + W - Wi) * 22).Value = iSmalGang
    InP.Cells(i + 2, 15 + (1 + W - Wi) * 22).Value = iLongGoods

    
    
  
    
'Getting data from HRM File-----------------------


    
    For iRow = 2 To LastRowHRM.Row
    
      If HRM.Cells(iRow, 13).Value = W Then 'W
    
        If arr(i) = HRM.Cells(iRow, 2).Value Then
            
            If HRM.Cells(iRow, 3).Value = 600 Or HRM.Cells(iRow, 3).Value = 604 Or HRM.Cells(iRow, 3).Value = 608 Or HRM.Cells(iRow, 3).Value = 617 Or HRM.Cells(iRow, 3).Value = 629 Or HRM.Cells(iRow, 3).Value = 630 Then
                iHRMOrdTruck = iHRMOrdTruck + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 601 Or HRM.Cells(iRow, 3).Value = 605 Or HRM.Cells(iRow, 3).Value = 609 Then
                iHRMHighLift = iHRMHighLift + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 603 Or HRM.Cells(iRow, 3).Value = 607 Or HRM.Cells(iRow, 3).Value = 611 Then
                iHRMElevator = iHRMElevator + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 602 Or HRM.Cells(iRow, 3).Value = 606 Or HRM.Cells(iRow, 3).Value = 618 Then
                iHRMSmalgang = iHRMSmalgang + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 616 Then
                iHRMLongGoods = iHRMLongGoods + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 628 Or HRM.Cells(iRow, 3).Value = 653 Then
                iHRMRepl = iHRMRepl + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value <> "" And HRM.Cells(iRow, 5).Value <> "RAST" Then
                iHRMOthers = iHRMOthers + HRM.Cells(iRow, 11).Value
            End If
        End If
     End If 'W
    Next
    
        

On Error Resume Next

'Registering Hours and productivity
    'Hours

    InP.Cells(i + 2, 18 + (1 + W - Wi) * 22).Value = iHRMRepl
    InP.Cells(i + 2, 8 + (1 + W - Wi) * 22).Value = iHRMOrdTruck
    InP.Cells(i + 2, 10 + (1 + W - Wi) * 22).Value = iHRMHighLift
    InP.Cells(i + 2, 12 + (1 + W - Wi) * 22).Value = iHRMElevator
    InP.Cells(i + 2, 14 + (1 + W - Wi) * 22).Value = iHRMSmalgang
    InP.Cells(i + 2, 16 + (1 + W - Wi) * 22).Value = iHRMLongGoods
    'Total Hours
    InP.Cells(i + 2, 19 + (1 + W - Wi) * 22).Value = (iHRMOrdTruck + iHRMHighLift + iHRMElevator + iHRMSmalgang + iHRMLongGoods)
    InP.Cells(i + 2, 20 + (1 + W - Wi) * 22).Value = iHRMRepl
    InP.Cells(i + 2, 21 + (1 + W - Wi) * 22).Value = iHRMOthers
    
    'Productivity
    InP.Cells(i + 2, 22 + (1 + W - Wi) * 22).Value = (iOrdTruck + iHighLift + iSmalGang + iLongGoods + iPaternost) / (iHRMOrdTruck + iHRMHighLift + iHRMElevator + iHRMSmalgang + iHRMLongGoods)
    InP.Cells(i + 2, 27 + (1 + W - Wi) * 22).Value = iRepl / iHRMRepl
    InP.Cells(i + 2, 23 + (1 + W - Wi) * 22).Value = iOrdTruck / iHRMOrdTruck
    InP.Cells(i + 2, 24 + (1 + W - Wi) * 22).Value = iHighLift / iHRMHighLift
    InP.Cells(i + 2, 25 + (1 + W - Wi) * 22).Value = iSmalGang / iHRMSmalgang
    InP.Cells(i + 2, 26 + (1 + W - Wi) * 22).Value = iLongGoods / iHRMLongGoods
    
    
    If InP.Cells(i + 2, 12 + (1 + W - Wi) * 22).Value > 0 Then InP.Cells(i + 2, 12 + (1 + W - Wi) * 22).Interior.ColorIndex = 50
    
 Next W
 
Next i

'Check cells  where there are missing HRM
Dim G As Long
For G = Wi To Wf
    For j = 8 To 18 Step 2
        For i = 3 To LastRowInP.Row

        If InP.Cells(i, j - 1 + (G - Wi) * 22).Value > 0 And InP.Cells(i, j + (G - Wi) * 22).Value = 0 Then
            InP.Cells(i, j + (G - Wi) * 22).Value = "No HRM Info"
            InP.Cells(i, j + (G - Wi) * 22).Interior.ColorIndex = 44
        End If
        Next
    Next j
Next G


End Sub


Sub Individual_EffectivityTotal()
'This calculates the individual effectivity for the total period of extraction

Dim InP As Worksheet
Set InP = Worksheets("Individual Performance")

Dim PRL As Worksheet
Set PRL = Worksheets("P&R Lines")

Dim HRM As Worksheet
Set HRM = Worksheets("HRM")

Dim LastRowInP As Range
Set LastRowInP = InP.Range("A900000").End(xlUp)




'Declaring SESAs as strings in the Individual sheet
Dim arr(1 To 160) As String

Dim K As Integer
For K = 3 To LastRowInP.Row
    arr(K - 2) = InP.Cells(K, 1).Value
Next



Dim i As Long
Set LastRowPL = PRL.Range("A900000").End(xlUp)
Set LastRowHRM = HRM.Range("A900000").End(xlUp)

'Dim of each variable
Dim iOrdTruck As Long
Dim iHighLift As Long
Dim iSmalGang As Long
Dim iLongGoods As Long
Dim iPaternost As Long
Dim iRepl As Long
Dim iTest As Long

Dim iR As Long

Dim iHRMOrdTruck As Double
Dim iHRMHighLift As Double
Dim iHRMElevator As Double
Dim iHRMSmalgang As Double
Dim iHRMLongGoods As Double
Dim iHRMRepl As Double



'Clean Sheet
InP.Range("F3:JI" & LastRowInP.Row).ClearContents
InP.Range("F3:JI" & LastRowInP.Row).Interior.Color = xlNone

'**********************************************************************************************************
'Get the total measured in the first set of data


'Loop items in array
For i = LBound(arr) To UBound(arr)

'Gathering of data in P&L Rep sheet------------------------------------------------------------
'Reset of variables


iOrdTruck = 0
iHighLift = 0
iSmalGang = 0
iLongGoods = 0
iPaternost = 0
iRepl = 0

iHRMOrdTruck = 0
iHRMHighLift = 0
iHRMElevator = 0
iHRMSmalgang = 0
iHRMElevator = 0
iHRMLongGoods = 0
iHRMRepl = 0
iHRMOthers = 0




    'PRL.Activate
    For iRow = 3 To LastRowPL.Row
    
    
        'Check if the line works for appointed operator
        If arr(i) = PRL.Cells(iRow, 17).Value Then

        
        
            If (PRL.Cells(iRow, 15).Value = 100 Or PRL.Cells(iRow, 15).Value = 916) And (PRL.Cells(iRow, 22).Value <> 20 Or PRL.Cells(iRow, 22).Value <> 21 Or PRL.Cells(iRow, 22).Value <> 120 Or PRL.Cells(iRow, 22).Value <> 121) Then
                'Count the total of picked Lines
                If (PRL.Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(PRL.Cells(iRow, 21).Value, 3) = "DPI" Or Left(PRL.Cells(iRow, 21).Value, 3) = "FBO" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAF" Then
                    iOrdTruck = iOrdTruck + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "ORD.ELKO") Or Left(PRL.Cells(iRow, 21).Value, 3) = "DPI" Or Left(PRL.Cells(iRow, 21).Value, 3) = "FBO" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAF" Then
                    iOrdTruck = iOrdTruck + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "HIGH LIFT") Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRP" Or Left(PRL.Cells(iRow, 21).Value, 3) = "HRF" Then
                    iHighLift = iHighLift + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "SMALGANG 1") Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAF" Then
                    iSmalGang = iSmalGang + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "SMALGANG_E") Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAD" Or Left(PRL.Cells(iRow, 21).Value, 3) = "NAF" Then
                    iSmalGang = iSmalGang + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "LONG GOODS") Then
                    iLongGoods = iLongGoods + 1
                ElseIf (PRL.Cells(iRow, 21).Value = "PATERNOST.") Or Left(PRL.Cells(iRow, 21).Value, 3) = "PAT" Then
                    iPaternost = iPaternost + 1
                End If
            ElseIf PRL.Cells(iRow, 21).Value = "REPL-HIGH" Or PRL.Cells(iRow, 21).Value = "REPL-LONG" Then
                iRepl = iRepl + 1
                
            End If
        
       End If

        
    Next iRow

    'Registering picking data for specific operator:
    InP.Cells(i + 2, 6).Value = iOrdTruck + iHighLift + iSmalGang + iLongGoods + iPaternost
    InP.Cells(i + 2, 17).Value = iRepl
    InP.Cells(i + 2, 7).Value = iOrdTruck
    InP.Cells(i + 2, 9).Value = iHighLift
    InP.Cells(i + 2, 11).Value = iPaternost
    InP.Cells(i + 2, 13).Value = iSmalGang
    InP.Cells(i + 2, 15).Value = iLongGoods

     
    
'Getting data from HRM File-----------------------


    
    For iRow = 2 To LastRowHRM.Row
    

    
        If arr(i) = HRM.Cells(iRow, 2).Value Then
            
            If HRM.Cells(iRow, 3).Value = 600 Or HRM.Cells(iRow, 3).Value = 604 Or HRM.Cells(iRow, 3).Value = 608 Or HRM.Cells(iRow, 3).Value = 617 Or HRM.Cells(iRow, 3).Value = 629 Or HRM.Cells(iRow, 3).Value = 630 Then
                iHRMOrdTruck = iHRMOrdTruck + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 601 Or HRM.Cells(iRow, 3).Value = 605 Or HRM.Cells(iRow, 3).Value = 609 Then
                iHRMHighLift = iHRMHighLift + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 603 Or HRM.Cells(iRow, 3).Value = 607 Or HRM.Cells(iRow, 3).Value = 611 Then
                iHRMElevator = iHRMElevator + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 602 Or HRM.Cells(iRow, 3).Value = 606 Or HRM.Cells(iRow, 3).Value = 618 Then
                iHRMSmalgang = iHRMSmalgang + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 616 Then
                iHRMLongGoods = iHRMLongGoods + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value = 628 Or HRM.Cells(iRow, 3).Value = 653 Then
                iHRMRepl = iHRMRepl + HRM.Cells(iRow, 11).Value
            ElseIf HRM.Cells(iRow, 3).Value <> "" And HRM.Cells(iRow, 5).Value <> "RAST" Then
                iHRMOthers = iHRMOthers + HRM.Cells(iRow, 11).Value
            End If
        End If
    Next
    
        

On Error Resume Next

'Registering Hours and productivity
    'Hours

    InP.Cells(i + 2, 18).Value = iHRMRepl
    InP.Cells(i + 2, 8).Value = iHRMOrdTruck
    InP.Cells(i + 2, 10).Value = iHRMHighLift
    InP.Cells(i + 2, 12).Value = iHRMElevator
    InP.Cells(i + 2, 14).Value = iHRMSmalgang
    InP.Cells(i + 2, 16).Value = iHRMLongGoods
    'Total Hours
    InP.Cells(i + 2, 19).Value = (iHRMOrdTruck + iHRMHighLift + iHRMElevator + iHRMSmalgang + iHRMLongGoods)
    InP.Cells(i + 2, 20).Value = iHRMRepl
    InP.Cells(i + 2, 21).Value = iHRMOthers
    
    'Productivity
    InP.Cells(i + 2, 22).Value = (iOrdTruck + iHighLift + iSmalGang + iLongGoods + iPaternost) / (iHRMOrdTruck + iHRMHighLift + iHRMElevator + iHRMSmalgang + iHRMLongGoods)
    InP.Cells(i + 2, 27).Value = iRepl / iHRMRepl
    InP.Cells(i + 2, 23).Value = iOrdTruck / iHRMOrdTruck
    InP.Cells(i + 2, 24).Value = iHighLift / iHRMHighLift
    InP.Cells(i + 2, 25).Value = iSmalGang / iHRMSmalgang
    InP.Cells(i + 2, 26).Value = iLongGoods / iHRMLongGoods
    
    
    If InP.Cells(i + 2, 12).Value > 0 Then InP.Cells(i + 2, 12).Interior.ColorIndex = 50
    

 
'Check cells  where there are missing HRM

For j = 8 To 18 Step 2
    For iR = 3 To LastRowInP.Row
    
    If InP.Cells(iR, j - 1).Value > 0 And InP.Cells(iR, j).Value = 0 Then
        InP.Cells(iR, j).Value = "No HRM Info"
        InP.Cells(iR, j).Interior.ColorIndex = 44
    End If
    Next
Next


Next i



End Sub

