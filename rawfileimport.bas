Attribute VB_Name = "Module10"
 Sub wmsstock_rawfileimport()
'
' wmsstock_raffileimport Makró
'

'

Dim filename As String
Dim deffilename As String
Dim wmstaskpath As String
Dim mainfilename As String

deffilename = Range("A9").Value
wmsstockpath = Range("A7").Value
mainfilename = Range("C1").Value

Sheets("WMS-Stock").Select
Rows("3:100000").Select
Selection.ClearContents
Range("A3").Select

filename = VBA.FileSystem.Dir(wmsstockpath & deffilename)
If filename = VBA.Constants.vbNullString Then
    MsgBox "File " & deffilename & " does not exist."

Else

Workbooks.Open wmsstockpath & deffilename, Format:=2, Delimiter:=","

'Workbooks.OpenText wmsstockpath & deffilename, DataType:=xlDelimited, Comma:=True, Local:=True

End If

Windows(deffilename).Activate

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("A2:Q2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows(mainfilename).Activate
Sheets("WMS-Stock").Select
Range("A3").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows(deffilename).Activate
Application.CutCopyMode = False
ActiveWindow.Close

MsgBox "Import of " & wmsstockpath & deffilename & " was successful."

End Sub
