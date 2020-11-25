Attribute VB_Name = "Module2"
Sub wmsstock_catformat()
Attribute wmsstock_catformat.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' catformat Makró
'
' Billentyûparancs: Ctrl+Shift+C
'
    Sheets("WMS-stock").Select
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("I3:I" & lastrow).Select
    Selection.TextToColumns Destination:=Range("T3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=">", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    
    Sheets("Dashboard").Select
    Range("B15").Select
    
End Sub
