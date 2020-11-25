Attribute VB_Name = "Module4"
Sub wmsstock_savefile()
Attribute wmsstock_savefile.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' savefile Makró
'
' Billentyûparancs: Ctrl+Shift+S
'

Sheets("WMS-stock").Select

Dim datum As String
Dim path As String
    

    'MsgBox (datum)


    'Mentés helyének meghatározása
    Range("C1").Select
    path = Selection.Value
    
    'datum meghatártozása
    datum = Replace(Range("A1").Value, "/", "")
    datum = Replace(datum, ":", "")
    datum = Replace(datum, ".", "")
    datum = Replace(datum, " ", "")
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks.Add
    
        Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ChDir path
    ActiveWorkbook.SaveAs filename:= _
        path & "WMS-Stock-" & datum & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close

    Range("A1").Select
    
    MsgBox "File saved to the following destination: " & path
    
End Sub
