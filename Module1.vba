Sub Macro1()
    '
    ' Macro1 Macro
    '
    
    '
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
    'Queue Group'!RC1,#REF!,"">0"",#REF!,""601"",#REF!,'Queue Group'!R1C)"
    Range("C2").Select
End Sub
Sub Macro2()
    '
    ' Macro2 Macro
    '
    
    '
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    Range("E3").Select
End Sub