Attribute VB_Name = "Module3"
Sub wmsstock_m3calculation()
Attribute wmsstock_m3calculation.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' m3calculation Makró
'
' Billentyûparancs: Ctrl+Shift+M
'
    Sheets("WMS-stock").Select
    Range("R3").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-3]*RC[-2]/1000000000"
    Range("R3").Select
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Selection.AutoFill Destination:=Range("R3:R" & lastrow)
    Range("R3:R" & lastrow).Select
    Selection.NumberFormat = "0.0000"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-13]"
    Selection.AutoFill Destination:=Range("S3:S" & lastrow)
    Range("S3:S" & lastrow).Select
    Selection.NumberFormat = "0.0000"
    
    Range("L3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0"
    
    Sheets("Dashboard").Select
    Range("A13").Select

End Sub
