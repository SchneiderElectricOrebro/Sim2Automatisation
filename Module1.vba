Sub Macro1()
    '
    ' Macro1 Macro
    '
    
    ' Select cell C2
    Range("C2").Select
    
    ' Set the formula in cell C2
    ActiveCell.FormulaR1C1 = _
    'Queue Group'!RC1,#REF!,"">0"",#REF!,""601"",#REF!,'Queue Group'!R1C)"
    
    ' Select cell C2 again (redundant, already selected)
    Range("C2").Select
End Sub

Sub Macro2()
    '
    ' Macro2 Macro
    '
    
    ' Select cell C2
    Range("C2").Select
    
    ' Set the formula in cell C2
    ActiveCell.FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    
    ' Select cell D2
    Range("D2").Select
    
    ' Set the formula in cell D2
    ActiveCell.FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    
    ' Select cell E2
    Range("E2").Select
    
    ' Set the formula in cell E2
    ActiveCell.FormulaR1C1 = _
    'P&R Lines'!C21,'Queue Group'!RC1,'P&R Lines'!C6,"">0"",'P&R Lines'!C11,""601"",'P&R Lines'!C25,'Queue Group'!R1C)"
    
    ' Select cell E3
    Range("E3").Select
End Sub