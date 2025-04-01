
Public Sub Monday()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A90000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & "." & Minute(Cells(iRow, 19))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) > "22:30") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) < "22:30") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) < "22:30") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
        'Condition if timeset matches what is being generated
    For iRow = LastRowPL.Row To 2 Step -1
            If (Cells(iRow, 24).Value = 1 And (Cells(iRow, 19).Value) > "22:30") Or (Cells(iRow, 24).Value = 2 And (Cells(iRow, 19).Value) < "23:00") Then
                'do nothing
            Else
                Rows(iRow).EntireRow.Delete
            End If
    Next
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
            ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + E
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If
                    
        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            
        End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
     
   'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If
        
    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    
    
    For iRow = 2 To LastRowHRM.Row
    
        If ((Cells(iRow, 7).Value) = "23:00") And (Cells(iRow, 12).Value) = 2 Then
        'do nothing
        Else
            If ((Cells(iRow, 7).Value) >= "22:30") Or ((Cells(iRow, 7).Value) < "06:00") Then
                Cells(iRow, 10).Value = "A3"
            Else
                If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                    If ((Cells(iRow, 7).Value) >= "14:30") And ((Cells(iRow, 7).Value) < "22:30") Then
                    Cells(iRow, 10).Value = "A2"
                    ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                    Cells(iRow, 10).Value = "A1"
                    End If
                ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                    If ((Cells(iRow, 7).Value) > "14:45") And ((Cells(iRow, 7).Value) < "22:30") Then
                    Cells(iRow, 10).Value = "A1"
                    ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                    Cells(iRow, 10).Value = "A2"
                    End If
                End If
            End If
        End If

        
        
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        
        'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    Next

        'condition if timeset is valid for day select
        For iRow = LastRowHRM.Row To 2 Step -1
            If (Cells(iRow, 12).Value = 1 And (Hour(Cells(iRow, 7)) >= 22)) Or (Cells(iRow, 12).Value = 2 And (Hour(Cells(iRow, 7))) < 23) Then
                'do nothing
            
            Else
                Rows(iRow).EntireRow.Delete
            End If
        Next

    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
        Next
        

 'Check Hours for Inbound
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If
        Next


Call Consolidation


End Sub

Public Sub Tuesday()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A900000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & "." & ((Minute(Cells(iRow, 19)) * 10 / 6))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) > "23:00") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) >= "14:45") And ((Cells(iRow, 19).Value) <= "23:00") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) <= "23:00") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
        'Condition if timeset matches what is being generated
    For iRow = LastRowPL.Row To 2 Step -1
            If (Cells(iRow, 24).Value = 2 And (Cells(iRow, 19).Value) >= "23:00") Or (Cells(iRow, 24).Value = 3 And (Cells(iRow, 19).Value) < "23:00") Then
                'do nothing
            Else
                Rows(iRow).EntireRow.Delete
            End If
    Next
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If
                    
        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            
        End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
    
  'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If
        
    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    
    
    For iRow = 2 To LastRowHRM.Row
        
        If ((Cells(iRow, 7).Value) > "23:00") Or ((Cells(iRow, 7).Value) < "06:00") Then
            Cells(iRow, 10).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                If ((Cells(iRow, 7).Value) >= "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A2"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                If ((Cells(iRow, 7).Value) > "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A1"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A2"
                End If
            End If
        End If
        
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        
       'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    Next

        'condition if timeset is valid for day select
        For iRow = LastRowHRM.Row To 2 Step -1
            If (Cells(iRow, 12).Value = 2 And (Hour(Cells(iRow, 7)) >= 23)) Or (Cells(iRow, 12).Value = 3 And (Hour(Cells(iRow, 7))) < 23) Then
                'do nothing
            
            Else
                Rows(iRow).EntireRow.Delete
            End If
        Next

    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
            
        Next
 
 'Check Hours for Inbound
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If

        Next
 
Call Consolidation



End Sub



Public Sub Wednesday()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A900000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & ":" & ((Minute(Cells(iRow, 19))))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) > "23:00") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) < "23:00") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) < "23:00") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
        'Condition if timeset matches what is being generated
    For iRow = LastRowPL.Row To 2 Step -1
            If (Cells(iRow, 24).Value = 3 And (Cells(iRow, 19).Value) >= "23:00") Or (Cells(iRow, 24).Value = 4 And (Cells(iRow, 19).Value) < "23:00") Then
                'do nothing
            Else
                Rows(iRow).EntireRow.Delete
            End If
    Next
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If

        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            
        End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
    
    'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If
    
    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    
    
    For iRow = 2 To LastRowHRM.Row
    
    
    
            If ((Cells(iRow, 7).Value) > "23:00") Or ((Cells(iRow, 7).Value) < "06:00") Then
            Cells(iRow, 10).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                If ((Cells(iRow, 7).Value) >= "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A2"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                If ((Cells(iRow, 7).Value) > "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A1"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A2"
                End If
            End If
        End If
        
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        
        'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    Next

        'condition if timeset is valid for day select
        For iRow = LastRowHRM.Row To 2 Step -1
            If (Cells(iRow, 12).Value = 3 And (Hour(Cells(iRow, 7)) >= 23)) Or (Cells(iRow, 12).Value = 4 And (Hour(Cells(iRow, 7))) < 23) Then
                'do nothing
            
            Else
                Rows(iRow).EntireRow.Delete
            End If
        Next

    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
        Next
        
 'Check Hours for Inbound
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If
        Next




Call Consolidation


End Sub


Public Sub Thursday()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A900000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & "." & Minute(Cells(iRow, 19))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) > "23:00") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) >= "14:45") And ((Cells(iRow, 19).Value) <= "23:00") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) >= "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) <= "23:00") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
        'Condition if timeset matches what is being generated
    For iRow = LastRowPL.Row To 2 Step -1
            If (Cells(iRow, 24).Value = 4 And (Cells(iRow, 19).Value) >= "23:00") Or (Cells(iRow, 24).Value = 5 And (Cells(iRow, 19).Value) < "23:00") Then
                'do nothing
            Else
                Rows(iRow).EntireRow.Delete
            End If
    Next
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If
                    
        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            
        End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
    
   'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If
    
    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    
    
    For iRow = 2 To LastRowHRM.Row
        
        
         If ((Cells(iRow, 7).Value) > "23:00") Or ((Cells(iRow, 7).Value) < "06:00") Then
            Cells(iRow, 10).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                If ((Cells(iRow, 7).Value) >= "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A2"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                If ((Cells(iRow, 7).Value) > "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A1"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A2"
                End If
            End If
        End If
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        
        'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    Next

        'condition if timeset is valid for day select
        For iRow = LastRowHRM.Row To 2 Step -1
            If (Cells(iRow, 12).Value = 4 And (Hour(Cells(iRow, 7)) >= 23)) Or (Cells(iRow, 12).Value = 5 And (Hour(Cells(iRow, 7))) < 23) Then
                'do nothing
            
            Else
                Rows(iRow).EntireRow.Delete
            End If
        Next

    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
        Next
        
 'Check Hours for Inbound
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If
        Next


Call Consolidation




End Sub


Public Sub Friday()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A900000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    Cells(1, 27).Value = "Location"
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & "." & Minute(Cells(iRow, 19))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) >= "23:00") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) >= "13:00") And ((Cells(iRow, 19).Value) < "21:00") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) >= "06:00") And ((Cells(iRow, 19).Value) < "13:00") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) >= "13:00") And ((Cells(iRow, 19).Value) < "21:00") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) >= "06:00") And ((Cells(iRow, 19).Value) < "13:00") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
        'Condition if timeset matches what is being generated
    For iRow = LastRowPL.Row To 2 Step -1
            If (Cells(iRow, 24).Value = 5 And (Cells(iRow, 19).Value) >= "23:00") Or (Cells(iRow, 24).Value = 6 And (Cells(iRow, 19).Value) < "21:00") Then
                'do nothing
            Else
                Rows(iRow).EntireRow.Delete
            End If
    Next
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If
                                      
        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            

    End If
    
        'End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
  
  'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If

    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    
    
    For iRow = 2 To LastRowHRM.Row
        
        If ((Cells(iRow, 7).Value) > "23:00") Or ((Cells(iRow, 7).Value) < "06:00") Then
            Cells(iRow, 10).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                If ((Cells(iRow, 7).Value) >= "13:00") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A2"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "13:00") Then
                Cells(iRow, 10).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                If ((Cells(iRow, 7).Value) > "13:00") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A1"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "13:00") Then
                Cells(iRow, 10).Value = "A2"
                End If
            End If
        End If
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    Next

        'condition if timeset is valid for day select
        For iRow = LastRowHRM.Row To 2 Step -1
            If (Cells(iRow, 12).Value = 5 And (Hour(Cells(iRow, 7)) >= 23)) Or (Cells(iRow, 12).Value = 6 And (Hour(Cells(iRow, 7))) < 21) Then
                'do nothing
            
            Else
                Rows(iRow).EntireRow.Delete
            End If
        Next

    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
        Next
        

 'Check Hours for Inbound
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If
        Next



Call Consolidation

End Sub

Public Sub Weekend()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A900000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & "." & Minute(Cells(iRow, 19))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) > "23:00") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) >= "14:45") And ((Cells(iRow, 19).Value) < "23:00") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) >= "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) < "23:00") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
        'Condition if timeset matches what is being generated
    For iRow = LastRowPL.Row To 2 Step -1
            If (Cells(iRow, 24).Value = 6 And (Cells(iRow, 19).Value) >= "21:30") Or (Cells(iRow, 24).Value = 7) Or (Cells(iRow, 24).Value = 1 And (Cells(iRow, 19).Value) < "22:00") Then
                'do nothing
            Else
                Rows(iRow).EntireRow.Delete
            End If
    Next
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If
                    
        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            
        End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
    
    'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If
    
    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    
    
    For iRow = 2 To LastRowHRM.Row
        
        
        
        If ((Cells(iRow, 7).Value) >= "23:00") Or ((Cells(iRow, 7).Value) < "06:00") Then
            Cells(iRow, 10).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                If ((Cells(iRow, 7).Value) >= "14:45") And ((Cells(iRow, 7).Value) < "23:00") Then
                Cells(iRow, 10).Value = "A2"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                If ((Cells(iRow, 7).Value) > "14:45") And ((Cells(iRow, 7).Value) < "23:00") Then
                Cells(iRow, 10).Value = "A1"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A2"
                End If
            End If
        End If
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        
        'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    Next

        'condition if timeset is valid for day select
        For iRow = LastRowHRM.Row To 2 Step -1
            If (Cells(iRow, 12).Value = 6 And (Hour(Cells(iRow, 7)) >= 22)) Or (Cells(iRow, 12).Value = 7) Or (Cells(iRow, 12).Value = 1 And (Hour(Cells(iRow, 7))) < 22) Then
                'do nothing
            
            Else
                Rows(iRow).EntireRow.Delete
            End If
        Next

    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
        Next
        
 'Check Hours for Inbound
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If
        Next

Call Consolidation

End Sub

Public Sub WholeWeek()

Sheets("P&R Lines").Select
    Set LastRowPL = Range("A900000").End(xlUp)
    
    Cells(1, 23).Value = "Exact Hour confirmed"
    Cells(1, 24).Value = "Week day"
    Cells(1, 25).Value = "Shift"
    Cells(1, 26).Value = "Week Number"
    
    
    For iRow = 2 To LastRowPL.Row
        Cells(iRow, 23).Value = Hour(Cells(iRow, 19)) & "." & Minute(Cells(iRow, 19))
        Cells(iRow, 24).Value = Weekday(Cells(iRow, 18))
        Cells(iRow, 26).Value = Application.WeekNum(Cells(iRow, 18))
               
        
        
        'adding row for Shift
        If ((Cells(iRow, 19).Value) > "23:00") Or ((Cells(iRow, 19).Value) < "06:00") Then
            Cells(iRow, 25).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = False Then
                If ((Cells(iRow, 19).Value) >= "14:45") And ((Cells(iRow, 19).Value) < "23:00") Then
                Cells(iRow, 25).Value = "A2"
                ElseIf ((Cells(iRow, 19).Value) >= "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 18))) = True Then
                If ((Cells(iRow, 19).Value) > "14:45") And ((Cells(iRow, 19).Value) < "23:00") Then
                Cells(iRow, 25).Value = "A1"
                ElseIf ((Cells(iRow, 19).Value) > "06:00") And ((Cells(iRow, 19).Value) < "14:45") Then
                Cells(iRow, 25).Value = "A2"
                End If
            End If
                
            
        End If
    Next
    
    
    
    
    'Now the verification of required data:
    
    
    Set LastRowPL = Range("A900000").End(xlUp)
    
    
    For iRow = 2 To LastRowPL.Row
        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            'Count the total of picked Lines
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then
                iOrdTruck = iOrdTruck + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iOrdTruckA1 = iOrdTruckA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iOrdTruckA2 = iOrdTruckA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iOrdTruckA3 = iOrdTruckA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then
                iHighLift = iHighLift + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iHighLiftA1 = iHighLiftA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iHighLiftA2 = iHighLiftA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iHighLiftA3 = iHighLiftA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
                    
                    ElseIf (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then
                iSmalGang = iSmalGang + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iSmalGangA1 = iSmalGangA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iSmalGangA2 = iSmalGangA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iSmalGangA3 = iSmalGangA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "LONG GOODS") Then
                iLongGoods = iLongGoods + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iLongGoodsA1 = iLongGoodsA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iLongGoodsA2 = iLongGoodsA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iLongGoodsA3 = iLongGoodsA3 + 1
                    End If
            
            ElseIf (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then
                iPaternost = iPaternost + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iPaternostA1 = iPaternostA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iPaternostA2 = iPaternostA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iPaternostA3 = iPaternostA3 + 1
                    End If
                    
            End If
        ElseIf Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then
            iRepl = iRepl + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iReplA1 = iReplA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iReplA2 = iReplA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iReplA3 = iReplA3 + 1
                    End If
                    
        'Count lines for Inbound
        ElseIf (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then
             iInbo = iInbo + 1
                    If (Cells(iRow, 25).Value = "A1") Then
                        iInboA1 = iInboA1 + 1
                    ElseIf (Cells(iRow, 25).Value = "A2") Then
                        iInboA2 = iInboA2 + 1
                    ElseIf (Cells(iRow, 25).Value = "A3") Then
                        iInboA3 = iInboA3 + 1
                    End If
            
        End If
    
     Set LastRowPL = Range("A90000").End(xlUp)
    
     Cells(1, 26).Value = "Hour"
      
     Cells(iRow, 26).Value = Hour(Cells(iRow, 19))
    
   'Add location for hourly follow up
Cells(1, 27).Value = "Location"

        If (Cells(iRow, 6).Value > 0 And Cells(iRow, 15).Value = 916) And (Cells(iRow, 22).Value <> 20 Or Cells(iRow, 22).Value <> 21 Or Cells(iRow, 22).Value <> 120 Or Cells(iRow, 22).Value <> 121) Then
            If (Cells(iRow, 21).Value = "ORD.TRUCK") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "ORD.ELKO") Or Left(Cells(iRow, 21).Value, 3) = "DPI" Or Left(Cells(iRow, 21).Value, 3) = "FBO" Or Left(Cells(iRow, 21).Value, 3) = "PAD" Or Left(Cells(iRow, 21).Value, 3) = "PAF" Then Cells(iRow, 27).Value = "Picking truck"
            If (Cells(iRow, 21).Value = "HIGH LIFT") Or Left(Cells(iRow, 21).Value, 3) = "HRD" Or Left(Cells(iRow, 21).Value, 3) = "HRP" Or Left(Cells(iRow, 21).Value, 3) = "HRF" Then Cells(iRow, 27).Value = "Reach truck"
            If (Cells(iRow, 21).Value = "SMALGANG 1") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "SMALGANG_E") Or Left(Cells(iRow, 21).Value, 3) = "NAD" Or Left(Cells(iRow, 21).Value, 3) = "NAF" Then Cells(iRow, 27).Value = "Narrow aisle"
            If (Cells(iRow, 21).Value = "LONG GOODS") Then Cells(iRow, 27).Value = "Long goods"
            If (Cells(iRow, 21).Value = "iPaternost") Or Left(Cells(iRow, 21).Value, 3) = "PAT" Then Cells(iRow, 27).Value = "Elevator"
            If Cells(iRow, 21).Value = "REPL-HIGH" Or Cells(iRow, 21).Value = "REPL-LONG" Then Cells(iRow, 27).Value = "Replenishment"
            If (Cells(iRow, 11).Value = 101 Or Cells(iRow, 11).Value = 105) And Cells(iRow, 6).Value > 0 Then Cells(iRow, 27).Value = "Inbound"
        End If
    
    Next
    

    
'Add columns for calculated time and week day and manipulation in HRM sheet-----------------------
    
    
    Sheets("HRM").Select
    Set LastRowHRM = Range("A900000").End(xlUp)
    
    Cells(1, 11).Value = "Time"
    Cells(1, 12).Value = "Week day"
    Cells(1, 13).Value = "Week Num"
    
    
    For iRow = 2 To LastRowHRM.Row
        
        
        If ((Cells(iRow, 7).Value) > "23:00") Or ((Cells(iRow, 7).Value) < "06:00") Then
            Cells(iRow, 10).Value = "A3"
        Else
            If Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = False Then
                If ((Cells(iRow, 7).Value) >= "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A2"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A1"
                End If
            ElseIf Application.IsEven(Application.WeekNum(Cells(iRow, 6))) = True Then
                If ((Cells(iRow, 7).Value) > "14:45") And ((Cells(iRow, 7).Value) <= "23:00") Then
                Cells(iRow, 10).Value = "A1"
                ElseIf ((Cells(iRow, 7).Value) > "06:00") And ((Cells(iRow, 7).Value) < "14:45") Then
                Cells(iRow, 10).Value = "A2"
                End If
            End If
        End If
        
        'calculating time worked
        If Cells(iRow, 7).Value < Cells(iRow, 8).Value Then
                Cells(iRow, 11).Value = (DateDiff("n", (Cells(iRow, 7)), (Cells(iRow, 8))) / 60)
            Else
                Cells(iRow, 11).Value = 24 - (DateDiff("n", (Cells(iRow, 8)), (Cells(iRow, 7))) / 60)
        End If
        
        
        'Addition of weekday corrected for night shift
        If (Hour(Cells(iRow, 7)) <= 5) Then
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6)) + 1
        Else
            Cells(iRow, 12).Value = Weekday(Cells(iRow, 6))
        End If
    
        'Week
        Cells(iRow, 13).Value = Application.WeekNum(Cells(iRow, 6))
    
    Next



    
    'Check Hours for Packing for each Shift
    Set LastRowHRM = Range("A90000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 10).Value = "A1" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA1 = iPackA1 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A2" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA2 = iPackA2 + Cells(iRow, 11).Value
            ElseIf Cells(iRow, 10).Value = "A3" And (Cells(iRow, 3).Value = 700 Or Cells(iRow, 3).Value = 701 Or Cells(iRow, 3).Value = 702 Or Cells(iRow, 3).Value = 703 Or Cells(iRow, 3).Value = 704 Or Cells(iRow, 3).Value = 705 Or Cells(iRow, 3).Value = 706 Or Cells(iRow, 3).Value = 707 Or Cells(iRow, 3).Value = 709 Or Cells(iRow, 3).Value = 711 Or Cells(iRow, 3).Value = 712 Or Cells(iRow, 3).Value = 713) Then
                iPackA3 = iPackA3 + Cells(iRow, 11).Value
            End If
        Next
        
 'Check Hours for Inbound
    Set LastRowHRM = Range("A900000").End(xlUp)
    
        For iRow = 2 To LastRowHRM.Row
            If Cells(iRow, 3).Value >= 500 And Cells(iRow, 3).Value <= 525 Then
                iInbo = iInbo + Cells(iRow, 11).Value
            End If
        Next

Call Consolidation

End Sub

