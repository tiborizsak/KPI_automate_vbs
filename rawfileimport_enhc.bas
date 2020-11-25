Attribute VB_Name = "Module11"
Sub wmsstock_task_import_consolidated()
Attribute wmsstock_task_import_consolidated.VB_ProcData.VB_Invoke_Func = " \n14"
'
' wmsstock_import_consolidated Makró
'

'

'VARIABLES

Dim stock_filename As String
Dim task_filename As String
Dim stock_deffilename As String
Dim task_deffilename As String
Dim stock_path As String
Dim task_path As String
Dim mainfilename As String
Dim stock_savepath As String
Dim task_savepath As String
Dim datum As String
Dim shortdatum As String
Dim sendto As String
Dim copyc As String
Dim subject As String
Dim body As String

Sheets("Dashboard").Select

stock_deffilename = Range("B11").Value
task_deffilename = Range("B12").Value
stock_path = Range("B5").Value
task_path = Range("B6").Value
mainfilename = Range("C1").Value
stock_savepath = Range("B8").Value
task_savepath = Range("B9").Value
datum = Range("F1").Value
    datum = Replace(datum, ":", "")
    datum = Replace(datum, ".", "")
    datum = Replace(datum, " ", "")
shortdatum = Range("G1").Value
    shortdatum = Replace(shortdatum, ".", "")
    
sendto = Range("B14").Value
copyc = Range("B16").Value
subject = Range("B18").Value
body = Range("B20").Value

'IMPORT MODULE

    'Clear content
Sheets("WMS-Stock").Select
Rows("3:100000").Select
Selection.ClearContents
Range("A3").Select

Sheets("WMS-Task").Select
Rows("3:100000").Select
Selection.ClearContents
Range("A3").Select

    'Check if STOCK file exists if yes import
filename = VBA.FileSystem.Dir(stock_path & stock_deffilename)
If filename = VBA.Constants.vbNullString Then
    MsgBox "File " & stock_deffilename & " does not exist."

Else

Workbooks.Open filename:=stock_path & stock_deffilename, Format:=2, Delimiter:=","

'Workbooks.Open filename:=stock_path & stock_deffilename, Local:=True

'Workbooks.Open stock_path & stock_deffilename, DataType:=xlDelimited, Comma:=True, Local:=True

End If

Windows(stock_deffilename).Activate

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("A2:Q2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows(mainfilename).Activate
Sheets("WMS-Stock").Select
Range("A3").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows(stock_deffilename).Activate
Application.CutCopyMode = False
ActiveWindow.Close

Windows(mainfilename).Activate

    'Check if TASK file exists if yes import
filename = VBA.FileSystem.Dir(task_path & task_deffilename)
If filename = VBA.Constants.vbNullString Then
    MsgBox "File " & task_deffilename & " does not exist."

Else

Workbooks.Open filename:=task_path & task_deffilename, Format:=2, Delimiter:=","

'Workbooks.Open filename:=stock_path & stock_deffilename, Format:=2, Delimiter:=","

'Workbooks.OpenText wmsstockpath & deffilename, DataType:=xlDelimited, Comma:=True, Local:=True

End If

Windows(task_deffilename).Activate

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("A2:Q2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows(mainfilename).Activate
Sheets("WMS-Task").Select
Range("A3").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows(task_deffilename).Activate
Application.CutCopyMode = False
ActiveWindow.Close

Windows(mainfilename).Activate

'IMPORT MODULE - END

'FORMAT MODULE

'Stock
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

'Catformat
Sheets("WMS-stock").Select
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("I3:I" & lastrow).Select
    Selection.TextToColumns Destination:=Range("T3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=">", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    
'Shippable
Range("AA3").Select
ActiveCell.FormulaR1C1 = "=IF(AND(RC[-2]<>""P1"",RC[-2]<>""HV"",RC[-7]<>""MDA""),""Y"",""N"")"
Selection.AutoFill Destination:=Range("AA3:AA" & lastrow)

Range("AA3:AA" & lastrow).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'M3
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

'Task
'Supracategory
Sheets("WMS-Task").Select
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Range("F3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "0"

'Name tasktype area
Range("F3").Select
Range(Selection, Selection.End(xlDown)).Select
ActiveWorkbook.Names.Add Name:="tasktype", RefersToR1C1:= _
    "='WMS-Task'!R3C6:R" & lastrow & "C6"

Range("R3").Select
ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP('WMS-Task'!RC[-12],'WMS-Stock'!R3C12:R1048576C20,9,0),""PickedorResealed"")"
    '"=INDEX('WMS-Stock'!RC[2]:R1048576C20,MATCH('WMS-Task'!RC[-12],'WMS-Stock'!R3C12:R1048576C12,0))"
Range("R3").Select
Selection.AutoFill Destination:=Range("R3:R" & lastrow)

Range("R3:R" & lastrow).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Application.CutCopyMode = False

'FORMAT MODULE - END

'SAVE MODULE
    'Stock

Sheets("WMS-Stock").Select
Range("A2").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks.Add

    Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

ChDir stock_savepath
ActiveWorkbook.SaveAs filename:= _
    stock_savepath & "WMS-Stock-" & datum & ".xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWindow.Close

    'Task
Sheets("WMS-Task").Select
Range("A2").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks.Add

    Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

ChDir task_savepath
ActiveWorkbook.SaveAs filename:= _
    task_savepath & "WMS-Task-" & datum & ".xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWindow.Close

Windows(mainfilename).Activate
Sheets("Dashboard").Select
Range("A1").Select

'MDA REPORT CREATION AND EMAIL

    Dim emailanswer As Integer
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    Sheets("WMS-Stock").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Range("A1").Select
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Munka1!R1C1:R" & lastrow & "C28", Version:=6).CreatePivotTable TableDestination:= _
        "Munka2!R3C1", TableName:="MDAPivot", DefaultVersion:=6
    Sheets("Munka2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("MDAPivot")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("MDAPivot").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("MDAPivot").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Zone")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Supra Cat")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Cat")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Sub")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("MDAPivot").AddDataField ActiveSheet.PivotTables( _
        "MDAPivot").PivotFields("tm3"), "Összeg / tm3", xlSum
    ActiveSheet.PivotTables("MDAPivot").AddDataField ActiveSheet.PivotTables( _
        "MDAPivot").PivotFields("Real qty"), "Összeg / Real qty", xlSum
  
    Columns("B:C").Select
    
    Selection.NumberFormat = "0.00"

    
    Range("A3").Select
    ActiveSheet.PivotTables("MDAPivot").RowGrand = False
    ActiveSheet.PivotTables("MDAPivot").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("MDAPivot").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Location").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Input Date").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Location type").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Block Reason").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Pal id").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Real qty").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Avail qty").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Max qty").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Supra").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Product").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("EIS id").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Part Number").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("default.weight (kg)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Width").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Height").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Lenght").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Parts").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("m3").Subtotals = Array(False _
        , False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("tm3").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Supra Cat").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Cat").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Sub").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Input date2").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Input Time").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Zone").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("RSG?").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Shippable?").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("GT?(2+ day)").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Zone").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Zone")
        '.PivotItems("AA").Visible = False
        .PivotItems("LB").Visible = False
        .PivotItems("LC").Visible = False
        .PivotItems("LR").Visible = False
        .PivotItems("MR").Visible = False
        .PivotItems("X.").Visible = False
    End With
    ActiveSheet.PivotTables("MDAPivot").PivotFields("Zone"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Supra Cat")
        .PivotItems("AC & Heating").Visible = False
        .PivotItems("Apparel Man").Visible = False
        .PivotItems("Audio, E-Books & Drones").Visible = False
        .PivotItems("AV & HiFi").Visible = False
        .PivotItems("Bags & Accessories").Visible = False
        .PivotItems("Books").Visible = False
        .PivotItems("Bricolage").Visible = False
        .PivotItems("Building materials").Visible = False
        .PivotItems("Car Accessories").Visible = False
        .PivotItems("Car Electronics").Visible = False
        .PivotItems("Car refrigerators").Visible = False
        .PivotItems("Care & Makeup").Visible = False
        .PivotItems("Children").Visible = False
        .PivotItems("Coffee and Tea").Visible = False
        .PivotItems("Components PC").Visible = False
        .PivotItems("Cycling").Visible = False
        .PivotItems("Desktop PCs").Visible = False
        .PivotItems("Detergents & Cleaners").Visible = False
        .PivotItems("Dining").Visible = False
        .PivotItems("Fitness equipment").Visible = False
        .PivotItems("Furniture and Matresses").Visible = False
        .PivotItems("Games & Gaming Acc").Visible = False
        .PivotItems("Gaming Consoles").Visible = False
        .PivotItems("Home Textiles").Visible = False
        .PivotItems("House Cleaning").Visible = False
        .PivotItems("Household").Visible = False
        .PivotItems("Hygiene").Visible = False
        .PivotItems("Laptops").Visible = False
        .PivotItems("Lighting & Electrical").Visible = False
    End With
    With ActiveSheet.PivotTables("MDAPivot").PivotFields("Supra Cat")
        .PivotItems("Luggages").Visible = False
        .PivotItems("MDA Others").Visible = False
        .PivotItems("Mobile Phones").Visible = False
        .PivotItems("Monitors").Visible = False
        .PivotItems("Office Supplies").Visible = False
        .PivotItems("Other Sports").Visible = False
        .PivotItems("Perfumes").Visible = False
        .PivotItems("Peripherals PC").Visible = False
        .PivotItems("Personal Care").Visible = False
        .PivotItems("Pet Shop").Visible = False
        .PivotItems("Phones Acc & Services").Visible = False
        .PivotItems("Photo & video accessories").Visible = False
        .PivotItems("Photo-Video").Visible = False
        .PivotItems("Printing Hardware").Visible = False
        .PivotItems("Printing Supplies").Visible = False
        .PivotItems("Sanitary").Visible = False
        .PivotItems("SDA").Visible = False
        .PivotItems("Season").Visible = False
        .PivotItems("Servers & Networking").Visible = False
        .PivotItems("Smart technology").Visible = False
        .PivotItems("Software").Visible = False
        .PivotItems("Sports clothing & footwear").Visible = False
        .PivotItems("Tablets").Visible = False
        .PivotItems("Tablets Acc & Services").Visible = False
        .PivotItems("Tires & Rims").Visible = False
        .PivotItems("Toys").Visible = False
        .PivotItems("TV").Visible = False
        .PivotItems("TV Acc").Visible = False
        .PivotItems("Vehicles").Visible = False
    End With
    
    Sheets("Munka2").Select
    Sheets("Munka2").Name = "Pivot"
    Sheets("Munka1").Select
    Sheets("Munka1").Name = "wms-stock-data"
    Sheets("Pivot").Select
    Sheets("Pivot").Move After:=Sheets(2)
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "Avg. M3"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "Avg. m3"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "Planned qty for inbound"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Planned total m3 inbound"
    Range("G3:I3").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A3:I3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("G3").Select
    With Selection.Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Range("H3:I3").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]/RC[-2]"
    Range("G4").Select
    Selection.AutoFill Destination:=Range("G4:G21"), Type:=xlFillDefault
    Range("G4:G21").Select
    
    Selection.NumberFormat = "0.00"
    Range("H4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC"
    Range("H4").Select
    Selection.ClearContents
    Range("I4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-2]"
    Range("I4").Select
    Selection.AutoFill Destination:=Range("I4:I21"), Type:=xlFillDefault
    Range("I4:I21").Select
    
    Selection.NumberFormat = "0.00"
    
    Range("G3:I21").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("G4:G21").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    Range("G22").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("G22").Select
    Selection.Font.Bold = True
    Range("H22").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-18]C:R[-1]C)"
    Range("I22").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-18]C:R[-1]C)"
    
    ChDir "\\WFS02-ALL\International$\HU\Warehouse\!Adminisztráció\KPI\MDA\archiv"
    ActiveWorkbook.SaveAs filename:= _
        "\\WFS02-ALL\International$\HU\Warehouse\!Adminisztráció\KPI\MDA\archiv\MDAriport-" & shortdatum & ".xlsb" _
        , FileFormat:=xlExcel12, CreateBackup:=False
        
    ActiveWindow.Close
        
      'KÉRDÉS ARRÓL, HOGY BIZTOSAN EL AKAROD-E KÜLDENI AZ MDA RIPORTOT
    
    emailanswer = MsgBox("Biztosan elküldöd az MDA Riportot?", vbQuestion + vbYesNo + vbDefaultButton2, "MDA RIPORT")

    If emailanswer = vbYes Then
        On Error Resume Next
        With OutMail
            .To = sendto
            .CC = copyc
            .BCC = ""
            .subject = subject
            .body = body
            '.Attachments.Add ActiveWorkbook.FullName
            'You can add other files also like this
            .Attachments.Add ("\\WFS02-ALL\International$\HU\Warehouse\!Adminisztráció\KPI\MDA\archiv\MDAriport-" & shortdatum & ".xlsb")
            .Send   'or use .Display
        End With
        On Error GoTo 0
    
        Set OutMail = Nothing
        Set OutApp = Nothing

    Else
        MsgBox "Az MDA riport nem lett elküldve"
    End If
    

'MDA REPORT CREATION AND EMAIL - END

Sheets("Output").Select

MsgBox "KÉSZ!"

End Sub
