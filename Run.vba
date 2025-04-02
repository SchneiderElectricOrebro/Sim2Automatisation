'Add columns for exact time and week day in PL Sheet-------------------------
Public iRow         As Long
Public LastRowPL    As Range
Public wPL          As Worksheet

Public iRowH        As Long
Public LastRowHRM   As Range

'Dim of each variable
Public iOrdTruck    As Long
Public iOrdTruckA1  As Long
Public iOrdTruckA2  As Long
Public iOrdTruckA3  As Long

Public iHighLift    As Long
Public iHighLiftA1  As Long
Public iHighLiftA2  As Long
Public iHighLiftA3  As Long

Public iSmalGang    As Long
Public iSmalGangA1  As Long
Public iSmalGangA2  As Long
Public iSmalGangA3  As Long

Public iLongGoods   As Long
Public iLongGoodsA1 As Long
Public iLongGoodsA2 As Long
Public iLongGoodsA3 As Long

Public iPaternost   As Long
Public iPaternostA1 As Long
Public iPaternostA2 As Long
Public iPaternostA3 As Long

Public iRepl        As Long
Public iReplA1      As Long
Public iReplA2      As Long
Public iReplA3      As Long

Public iPackA1      As Double
Public iPackA2      As Double
Public iPackA3      As Double

Public iInbo        As Long
Public iInboA1      As Long
Public iInboA2      As Long
Public iInboA3      As Long

'Set UltCelVert = wb1.Sheets(1).Range("A90000").End(xlUp)
Public Sub Run()
    
    'Dim of each variable
    iOrdTruck = 0
    iOrdTruckA1 = 0
    iOrdTruckA2 = 0
    iOrdTruckA3 = 0
    
    iHighLift = 0
    iHighLiftA1 = 0
    iHighLiftA2 = 0
    iHighLiftA3 = 0
    
    iSmalGang = 0
    iSmalGangA1 = 0
    iSmalGangA2 = 0
    iSmalGangA3 = 0
    
    iLongGoods = 0
    iLongGoodsA1 = 0
    iLongGoodsA2 = 0
    iLongGoodsA3 = 0
    
    iPaternost = 0
    iPaternostA1 = 0
    iPaternostA2 = 0
    iPaternostA3 = 0
    
    iRepl = 0
    iReplA1 = 0
    iReplA2 = 0
    iReplA3 = 0
    
    iPackA1 = 0
    iPackA2 = 0
    iPackA3 = 0
    
    iInbo = 0
    iInboA1 = 0
    iInboA2 = 0
    iInboA3 = 0
    
    Application.ScreenUpdating = FALSE
    
    'Clear previous data
    If Worksheets("HRM").AutoFilterMode Then Worksheets("HRM").AutoFilterMode = FALSE
    If Worksheets("P&R Lines").AutoFilterMode Then Worksheets("P&R Lines").AutoFilterMode = FALSE
    
    Sheets("Data").Range("B9:B14,B17:B20,E10:F17,I10:J17,M10:N17,Q10:S13").ClearContents
    
    'apply logic to erase data from earlier day according to present week day
    If Format(Now(), "DDD") = "Mon" Then
        'run monday code****************************************************************************************************************************************
        
        Call Friday
        
    End If
    
    If Format(Now(), "DDD") = "Tue" Then
        'run tuesday code
        '**********************************************************************************************************************************
        Call Monday
        
    End If
    
    If Format(Now(), "DDD") = "Wed" Then
        'run wednesday code********************************************************************************************************************************************
        Call Tuesday
        
    End If
    
    If Format(Now(), "DDD") = "Thu" Then
        'run thursday code*********************************************************************************************************************************************
        Call Wednesday
    End If
    
    If Format(Now(), "DDD") = "Fri" Then
        'run Friday code**********************************************************************************************************************************************
        Call Thursday
        
    End If
    
End Sub

Public Sub Consolidation()
    
    'Common Code for all shifts
    '=======================================================================================================
    
    'HRM Data
    Worksheets("Queue Group").Select
    Range("C2:C19").FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    Range("D2:D19").FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    Range("E2:E19").FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    
    Worksheets("Data").Select
    
    Range("B31").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(HRM!C[9],HRM!C[1],LEFT(RC[-1],3))"
    Selection.AutoFill Destination:=Range("B31:B154")
    
    Range("O31").Select
    ActiveCell.FormulaR1C1 = _
                             "=SUMIFS(HRM!C[-4],HRM!C[-12],LEFT(Data!RC[-1],3),HRM!C[-5],""A1"")"
    Selection.AutoFill Destination:=Range("O31:O77")
    
    Range("R31").Select
    ActiveCell.FormulaR1C1 = _
                             "=SUMIFS(HRM!C[-7],HRM!C[-15],LEFT(Data!RC[-1],3),HRM!C[-8],""A2"")"
    Selection.AutoFill Destination:=Range("R31:R77")
    
    Range("U31").Select
    ActiveCell.FormulaR1C1 = _
                             "=SUMIFS(HRM!C[-10],HRM!C[-18],LEFT(Data!RC[-1],3),HRM!C[-11],""A3"")"
    Selection.AutoFill Destination:=Range("U31:U77")
    
    'Other Tasks:
    Range("F19").Value = Application.Sum(Range("O45,O46,O50:O55,O57,O58,O63"))
    Range("J19").Value = Application.Sum(Range("R45,R46,R50:R55,R57,R58,R63"))
    Range("N19").Value = Application.Sum(Range("U45,U46,U50:U55,U57,U58,U63"))
    
    'Data for Pick and Repl
    Worksheets("Data").Range("B10").Value = iOrdTruck + iHighLift + iSmalGang + iLongGoods
    Worksheets("Data").Range("B11").Value = iRepl
    Worksheets("Data").Range("I31").Value = iOrdTruck
    Worksheets("Data").Range("I32").Value = iHighLift
    Worksheets("Data").Range("I33").Value = Range("D3").Value + Range("D4").Value + Range("D5").Value
    Worksheets("Data").Range("I34").Value = iSmalGang
    Worksheets("Data").Range("I35").Value = iLongGoods
    
    'A1
    Worksheets("Data").Range("F10").Value = iOrdTruckA1 + iHighLiftA1 + iSmalGangA1 + iLongGoodsA1 + Worksheets("Data").Range("D3").Value
    Worksheets("Data").Range("F11").Value = iReplA1
    Worksheets("Data").Range("F13").Value = iOrdTruckA1
    Worksheets("Data").Range("F14").Value = iHighLiftA1
    Worksheets("Data").Range("F15").Value = Worksheets("Data").Range("D3").Value
    Worksheets("Data").Range("F16").Value = iSmalGangA1
    Worksheets("Data").Range("F17").Value = iLongGoodsA1
    
    On Error Resume Next
    
    Worksheets("Data").Range("E10").Value = (iOrdTruckA1 + iHighLiftA1 + iSmalGangA1 + iLongGoodsA1 + Worksheets("Data").Range("D3").Value) / (Application.Sum(Range("O31:O45,O47:O58")) + Range("O63").Value)
    Worksheets("Data").Range("E11").Value = iReplA1 / (Range("O59").Value + Range("O62").Value)
    Worksheets("Data").Range("E13").Value = iOrdTruckA1 / (Range("O31").Value + Range("O35").Value + Range("O39").Value + Range("O43").Value + Range("O44").Value + Range("O48").Value)
    Worksheets("Data").Range("E14").Value = iHighLiftA1 / (Range("O32").Value + Range("O36").Value + Range("O40").Value)
    Worksheets("Data").Range("E15").Value = Worksheets("Data").Range("D3").Value / (Range("O34").Value + Range("O38").Value + Range("O42").Value)
    Worksheets("Data").Range("E16").Value = iSmalGangA1 / (Range("O33").Value + Range("O37").Value + Range("O49").Value)
    Worksheets("Data").Range("E17").Value = iLongGoodsA1 / Range("O47").Value
    
    'A2
    
    Worksheets("Data").Range("J10").Value = iOrdTruckA2 + iHighLiftA2 + iSmalGangA2 + iLongGoodsA2 + Worksheets("Data").Range("D4").Value
    Worksheets("Data").Range("J11").Value = iReplA2
    Worksheets("Data").Range("J13").Value = iOrdTruckA2
    Worksheets("Data").Range("J14").Value = iHighLiftA2
    Worksheets("Data").Range("J15").Value = Worksheets("Data").Range("D4").Value
    Worksheets("Data").Range("J16").Value = iSmalGangA2
    Worksheets("Data").Range("J17").Value = iLongGoodsA2
    
    Worksheets("Data").Range("I10").Value = (iOrdTruckA2 + iHighLiftA2 + iSmalGangA2 + iLongGoodsA2 + Worksheets("Data").Range("D4").Value) / (Application.Sum(Range("R31:R45,R47:R58")) + Range("R63").Value)
    Worksheets("Data").Range("I11").Value = iReplA2 / (Range("R59").Value + Range("R62").Value)
    Worksheets("Data").Range("I13").Value = iOrdTruckA2 / (Range("R31").Value + Range("R35").Value + Range("R39").Value + Range("R43").Value + Range("R44").Value + Range("R48").Value)
    Worksheets("Data").Range("I14").Value = iHighLiftA2 / (Range("R32").Value + Range("R36").Value + Range("R40").Value)
    Worksheets("Data").Range("I15").Value = Worksheets("Data").Range("D4").Value / (Range("R34").Value + Range("R38").Value + Range("R42").Value)
    Worksheets("Data").Range("I16").Value = iSmalGangA2 / (Range("R33").Value + Range("R37").Value + Range("R49").Value)
    Worksheets("Data").Range("I17").Value = iLongGoodsA2 / Range("R47").Value
    
    'A3
    Worksheets("Data").Range("N10").Value = iOrdTruckA3 + iHighLiftA3 + iSmalGangA3 + iLongGoodsA3 + Worksheets("Data").Range("D5").Value
    Worksheets("Data").Range("N11").Value = iReplA3
    Worksheets("Data").Range("N13").Value = iOrdTruckA3
    Worksheets("Data").Range("N14").Value = iHighLiftA3
    Worksheets("Data").Range("N15").Value = Worksheets("Data").Range("D5").Value
    Worksheets("Data").Range("N16").Value = iSmalGangA3
    Worksheets("Data").Range("N17").Value = iLongGoodsA3
    
    Worksheets("Data").Range("M10").Value = (iOrdTruckA3 + iHighLiftA3 + iSmalGangA3 + iLongGoodsA3 + Worksheets("Data").Range("D5").Value) / (Application.Sum(Range("U31:U45,U47:U58")) + Range("U63").Value)
    Worksheets("Data").Range("M11").Value = iReplA3 / (Range("U59").Value + Range("U62").Value)
    Worksheets("Data").Range("M13").Value = iOrdTruckA3 / (Range("U31").Value + Range("U35").Value + Range("U39").Value + Range("U43").Value + Range("U44").Value + Range("U48").Value)
    Worksheets("Data").Range("M14").Value = iHighLiftA3 / (Range("U32").Value + Range("U36").Value + Range("U40").Value)
    Worksheets("Data").Range("M15").Value = Worksheets("Data").Range("D5").Value / (Range("U34").Value + Range("U38").Value + Range("U42").Value)
    Worksheets("Data").Range("M16").Value = iSmalGangA3 / (Range("U33").Value + Range("U37").Value + Range("U49").Value)
    Worksheets("Data").Range("M17").Value = iLongGoodsA3 / Range("U47").Value
    
    'Packing of Staging Area 01
    
    Worksheets("Data").Range("Q15").Value = Range("F3").Value
    Worksheets("Data").Range("Q16").Value = Range("F4").Value
    Worksheets("Data").Range("Q17").Value = Range("F5").Value
    Worksheets("Data").Range("Q18").Value = Range("F3").Value + Range("F4").Value + Range("F5").Value
    
    Worksheets("Data").Range("R15").Value = Worksheets("Data").Range("O67").Value
    Worksheets("Data").Range("R16").Value = Worksheets("Data").Range("R67").Value
    Worksheets("Data").Range("R17").Value = Worksheets("Data").Range("U67").Value
    Worksheets("Data").Range("R18").Value = Worksheets("Data").Range("B120").Value
    
    Worksheets("Data").Range("S15").Value = Range("Q15").Value / Range("R15").Value
    Worksheets("Data").Range("S16").Value = Range("Q16").Value / Range("R16").Value
    Worksheets("Data").Range("S17").Value = Range("Q17").Value / Range("R17").Value
    Worksheets("Data").Range("S18").Value = Range("Q18").Value / Range("R18").Value
    
    'Data for Packing
    
    Worksheets("Data").Range("Q10").Value = Range("E3").Value - Range("F3").Value
    Worksheets("Data").Range("Q11").Value = Range("E4").Value - Range("F4").Value
    Worksheets("Data").Range("Q12").Value = Range("E5").Value - Range("F5").Value
    Worksheets("Data").Range("Q13").Value = Range("E3").Value + Range("E4").Value + Range("E5").Value - Range("Q18").Value
    
    Worksheets("Data").Range("R10").Value = iPackA1 - Range("R15").Value
    Worksheets("Data").Range("R11").Value = iPackA2 - Range("R16").Value
    Worksheets("Data").Range("R12").Value = iPackA3 - Range("R17").Value
    Worksheets("Data").Range("R13").Value = iPackA1 + iPackA2 + iPackA3 - Range("R18")
    
    Worksheets("Data").Range("S10").Value = Range("Q10").Value / Range("R10").Value
    Worksheets("Data").Range("S11").Value = Range("Q11").Value / Range("R11").Value
    Worksheets("Data").Range("S12").Value = Range("Q12").Value / Range("R12").Value
    Worksheets("Data").Range("S13").Value = Range("Q13").Value / Range("R13").Value
    
    Worksheets("Data").Range("R20").Value = Range("B124").Value
    
    'Data for Inbound
    Worksheets("Data").Range("V10").Value = iInboA1 + iInboA2 + iInboA3
    Range("W10").Value = Application.Sum(Range("B55:B80")) + Worksheets("Data").Range("B154").Value
    Worksheets("Data").Range("X10").Value = (iInboA1 + iInboA2 + iInboA3) / Range("W10")
    
    'Summary
    Worksheets("Data").Range("B09").Value = Application.Sum(Range("D3:D5")) + Worksheets("Data").Range("B10").Value
    Worksheets("Data").Range("B12").Value = Application.Sum(Range("B84:B111")) + Worksheets("Data").Range("B116").Value - Worksheets("Data").Range("B99").Value + Worksheets("Data").Range("B150").Value
    Worksheets("Data").Range("B13").Value = Worksheets("Data").Range("B12").Value - (Worksheets("Data").Range("B104").Value + Worksheets("Data").Range("B105").Value + Worksheets("Data").Range("B108").Value + Worksheets("Data").Range("B117").Value + Worksheets("Data").Range("B126").Value) - Worksheets("Data").Range("B99").Value
    Worksheets("Data").Range("B14").Value = Worksheets("Data").Range("B112").Value + Worksheets("Data").Range("B115").Value
    Worksheets("Data").Range("B22").Value = Worksheets("Data").Range("B99").Value
    
    Worksheets("Data").Range("B18").Value = Range("B09").Value / Range("L37").Value
    Worksheets("Data").Range("B19").Value = Range("B09").Value / Range("B12").Value
    Worksheets("Data").Range("B20").Value = Range("B09").Value / Range("B13").Value
    Worksheets("Data").Range("B17").Value = Range("B18").Value - Range("K40").Value
    
    Worksheets("Data").Range("E12").Value = Range("K56").Value
    Worksheets("Data").Range("I12").Value = Range("K69").Value
    Worksheets("Data").Range("M12").Value = Range("K82").Value
    
    'Show other hours for each shift
    Set Data = Worksheets("Data")
    
    Range("N45,N46,N50:N55,N57,N58,N63,Q45,Q46,Q50:Q55,Q57,Q58,Q63,T45,T46,T50:T55,T57,T58,T63").Font.Color = vbBlue
    Data.Range(Cells(20, 4), Cells(28, 14)).ClearContents
    
    For iRow = 45 To 63
        If Data.Cells(iRow, 15).Value > 0 And Data.Cells(iRow, 14).Font.Color = vbBlue Then
            Data.Cells(20 + iL, 4).Value = Data.Cells(iRow, 14).Value
            Data.Cells(20 + iL, 6).Value = Data.Cells(iRow, 15).Value
            iL = iL + 1
        End If
        
        If Data.Cells(iRow, 18).Value > 0 And Data.Cells(iRow, 17).Font.Color = vbBlue Then
            Data.Cells(20 + iH, 8).Value = Data.Cells(iRow, 17).Value
            Data.Cells(20 + iH, 10).Value = Data.Cells(iRow, 18).Value
            iH = iH + 1
        End If
        
        If Data.Cells(iRow, 21).Value > 0 And Data.Cells(iRow, 20).Font.Color = vbBlue Then
            Data.Cells(20 + iK, 12).Value = Data.Cells(iRow, 20).Value
            Data.Cells(20 + iK, 14).Value = Data.Cells(iRow, 21).Value
            iK = iK + 1
        End If
    Next
    
    Unload SelectDate
    Worksheets("Data").Select
    
    End
End Sub