Attribute VB_Name = "Module1"
Sub wmsstock_dateformat()
Attribute wmsstock_dateformat.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' dateformat Makró
'
' Billentyûparancs: Ctrl+Shift+D
'
    'Manipulation area
    Sheets("WMS-stock").Select
    Range("AI3").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-33],7,4)"
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Selection.AutoFill Destination:=Range("AI3:AI" & lastrow)
    Range("AJ3").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-34],4,2)"
    Selection.AutoFill Destination:=Range("AJ3:AJ" & lastrow)
    Range("AK3").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-35],2)"
    Selection.AutoFill Destination:=Range("AK3:AK" & lastrow)
    Range("W3").Select
    ActiveCell.FormulaR1C1 = "=DATE(RC[12],RC[13],RC[14])"
    Selection.AutoFill Destination:=Range("W3:W" & lastrow)
    Range("X3").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-22],5)"
    Selection.AutoFill Destination:=Range("X3:X" & lastrow)
    
    Range("W3:X" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Zone
    Range("Y3").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-24],2)"
    Selection.AutoFill Destination:=Range("Y3:Y" & lastrow)
    
    Range("Y3:Y" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'RSG
    Range("Z3").Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(RC[-14],7)=""RSGEMAG"",""Y"",""N"")"
    Selection.AutoFill Destination:=Range("Z3:Z" & lastrow)
    
    Range("Z3:Z" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Shippable
    Range("AA3").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-2]<>""P1"",RC[-2]<>""HV"",RC[-7]<>""MDA""),""Y"",""N"")"
    Selection.AutoFill Destination:=Range("AA3:AA" & lastrow)
    
    Range("AA3:AA" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'GT + 2
    Range("AB3").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-3]=""GT"",(R1C1-RC[-5])>2),""Y"",""N"")"
    Selection.AutoFill Destination:=Range("AB3:AB" & lastrow)
    
    Range("AB3:AB" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    'Name PN area
    Range("L3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Names.Add Name:="partnumber", RefersToR1C1:= _
        "='WMS-Stock'!R3C12:R" & lastrow & "C12"
    
    Range("AI3:AK" & lastrow).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    Range("A2").Select
    
    Sheets("Dashboard").Select
    Range("B16").Select
End Sub
