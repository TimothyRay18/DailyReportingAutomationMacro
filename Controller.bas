Attribute VB_Name = "Controller"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function

Sub addBorder(start As String, last As String)
    Range(start & ":" & last).Select
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
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
End Sub

Sub JIT()
'
' JIT Macro
'

'
    Dim wb As String
    wb = "JIT " + CStr(Format(Now(), "dd")) + " " + UCase(Format(Date, "mmmm")) + " " + CStr(Format(Now(), "yy"))
    Sheets("Sheet1").Name = wb
    Rows("1:8").Select
    Selection.Delete Shift:=xlUp
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Cells.Select
    Range("A2").Activate
    Selection.Columns.AutoFit
    
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("G:H").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("AB:AB").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("AC:AC").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("BI:BK").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("BJ:BK").Select
    Selection.Delete Shift:=xlToLeft
    
    Range(Columns(findCellInColumn(1, "Follow up material") + 1), Columns(findCellInColumn(1, "Standard Cost") - 1)).Select
    Selection.EntireColumn.Hidden = True
    
    unit_col = findCellInColumn(1, "Standard Cost") + 1
    Columns(unit_col + 1).Select
    Selection.Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Cells(1, (unit_col + 1)).Select
    ActiveCell.FormulaR1C1 = "Per Piece"
    
    Columns(unit_col + 1).EntireColumn.AutoFit
    
    Cells(2, (unit_col + 1)).Select
    Selection.NumberFormat = "#,##0.0000"
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]/RC[-1]"
    
    Selection.AutoFill Destination:=Range(Cells(2, (unit_col + 1)), Cells(getMaxRow(3), (unit_col + 1)))
    
    addBorder "A1", Cells(getMaxRow(3), getMaxCol(1)).Address
End Sub

Sub BF()
'
' BF Macro
'

'
    If IsEmpty(Range("A2").Value) = True Then
        Rows("1:3").Select
        Selection.Delete Shift:=xlUp
        
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
    End If
    
    Dim max_row As String
    max_row = CStr(getMaxRow(1))
    max_col = getMaxCol(1)
    Range("A1", Cells(max_row, CDbl(max_col))).Select
    Selection.Columns.AutoFit
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    If findCellInColumn(2, " Amount LC") < findCellInColumn(2, "Amount in LC") Then
        lc_col = findCellInColumn(2, " Amount LC")
    Else
        lc_col = findCellInColumn(2, "Amount in LC")
    End If
    
 '   Cells(1, findCellInColumn(2, " Amount LC")).Select
    Cells(1, lc_col).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" + max_row + "]C)"
'   Columns(findCellInColumn(2, " Amount LC")).Select
    Columns(lc_col).Select
    Selection.NumberFormat = _
        "_-[$$-en-US]* #,##0_ ;_-[$$-en-US]* -#,##0 ;_-[$$-en-US]* ""-""??_ ;_-@_ "
    
'    Cells(1, findCellInColumn(2, " Amount LC") + 1).Select
    Cells(1, lc_col + 1).Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/1000000"
    Selection.NumberFormat = _
        "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "
    
'    addBorder Cells(1, findCellInColumn(2, " Amount LC")).Address, Cells(1, findCellInColumn(2, " Amount LC") + 1).Address
    addBorder Cells(1, lc_col).Address, Cells(1, lc_col + 1).Address
    
    addBorder "A2", Cells(getMaxRow(3), getMaxCol(2)).Address
    
    Columns(findCellInColumn(2, "Material")).Select
    Selection.NumberFormat = "@"
End Sub

Sub FG()
'
' FG Macro
'

'
    If IsEmpty(Range("A2").Value) = True Then
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
        
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
        
        Columns("A:B").Select
        Selection.Delete Shift:=xlToLeft
        
        Dim r As String
        r = CStr(getMaxRow(1) + 1)
        Rows(r + ":" + r).Select
        Selection.Delete Shift:=xlUp
    End If
    
    max_row = getMaxRow(1)
    max_col = getMaxCol(1)
    Range("A1", Cells(max_row, max_col)).Select
    Selection.Columns.AutoFit
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Cells(2, max_col + 1).Select
    ActiveCell.FormulaR1C1 = "Total Stock"
    
    Cells(3, max_col + 1).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]+RC[-3]"
    
    Cells(3, max_col + 1).Select
    Selection.AutoFill Destination:=Range(Cells(3, max_col + 1), Cells(max_row + 1, max_col + 1))
    
    Range(Columns(max_col - 2), Columns(max_col + 1)).Select
    Selection.NumberFormat = _
        "_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* ""-""??_);_(@_)"
        
        
    Dim bottom As Double
    
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" + CStr(getMaxRow(12) - 2) + "]C)"
    
    Cells(1, max_col - 2).Select
    Selection.AutoFill Destination:=Range(Cells(1, max_col - 2), Cells(1, max_col + 1)), Type:=xlFillDefault
    addBorder Cells(1, max_col - 2).Address, Cells(1, max_col + 1).Address
    
    Range(Cells(1, max_col - 2), Cells(1, max_col + 1)).Select
    Columns(max_col + 1).EntireColumn.AutoFit
    
    addBorder "A2", Cells(max_row + 1, max_col + 1).Address
    
    Columns(findCellInColumn(2, "Material")).Select
    Selection.NumberFormat = "@"
End Sub

Sub DI()
'
' DI Macro
'

'
    If IsEmpty(Range("A2").Value) = True Then
        Rows("1:8").Select
        Selection.Delete Shift:=xlUp
        
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
        
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        
        Columns("C:E").Select
        Selection.Delete Shift:=xlToLeft
    End If
    
    Dim max_row As Double
    max_row = getMaxRow(1)
    Dim max_col As Double
    max_col = getMaxCol(2)
    
    Range(Cells(1, 1), Cells(max_row, max_col)).Select
    Selection.Columns.AutoFit
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Columns(findCellInColumn(2, "Short text") + 1).Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" + CStr(max_row) + "]C)"
    
    Dim st As Double
    st = findCellInColumn(2, "Short text")
    
    Range(Cells(2, st + 1), Cells(2, st + 2)).Select
    Selection.Merge
    
    Cells(1, st + 1).Select
    Selection.NumberFormat = _
        "_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* ""-""??_);_(@_)"
        
    addBorder Cells(1, st + 1).Address, Cells(1, st + 1).Address
    
    addBorder "A2", Cells(getMaxRow(1), getMaxCol(3)).Address
End Sub

Sub GR_template(t As String)
Attribute GR_template.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GR_template Macro
'

'
    If IsEmpty(Range("A2").Value) = True Then
        Rows("1:3").Select
        Selection.Delete Shift:=xlUp
        
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
    End If
    
    Dim max_row As Double
    max_row = getMaxRow(1)
    Dim max_col As Double
    max_col = getMaxCol(1)
    
    Range(Cells(1, 1), Cells(max_row, max_col)).Select
    Selection.Columns.AutoFit
    
    Cells(1, getMaxCol(1) + 1).Select
    ActiveCell.FormulaR1C1 = "Buyer"
    
    Cells(1, getMaxCol(1) + 1).Select
    ActiveCell.FormulaR1C1 = "Supplier"
    Cells(1, getMaxCol(1)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
    End With
    
    ThisWorkbook.Activate
    ActiveWorkbook.Worksheets("Sheet1").Activate
    
    Columns(findCellInColumn(1, "Vendor")).Select
    Selection.TextToColumns Destination:=Cells(1, findCellInColumn(1, "Vendor")), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    
    
    Cells(2, findCellInColumn(1, "Buyer")).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-" + CStr(findCellInColumn(1, "Buyer") - findCellInColumn(1, "Material")) + "],'Master Data'!C1:C2,2,0)"
        
    Cells(2, findCellInColumn(1, "Supplier")).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-" + CStr(findCellInColumn(1, "Supplier") - findCellInColumn(1, "Vendor")) + "],'Master Data'!C4:C5,2,0)"
        
    Range(Cells(2, findCellInColumn(1, "Buyer")), Cells(2, findCellInColumn(1, "Supplier"))).Select
    Selection.AutoFill Destination:=Range(Cells(2, findCellInColumn(1, "Buyer")), Cells(getMaxRow(1), findCellInColumn(1, "Supplier")))
    
    Range(Cells(1, 1), Cells(getMaxRow(1), findCellInColumn(1, "Supplier"))).Select
    Selection.AutoFilter
    
    Columns(findCellInColumn(1, "Supplier")).EntireColumn.AutoFit
    Columns(findCellInColumn(1, "Buyer")).EntireColumn.AutoFit
    
    addBorder "A1", Cells(getMaxRow(1), getMaxCol(1)).Address
    
    Dim c As String
    c = Cells(getMaxRow(1), getMaxCol(1)).Address
    
  '  RemoveFormula
    
    If t = 5 Then
        ActiveSheet.Range("$A$1:" + c).AutoFilter Field:=4, Criteria1:="<>"
    End If
    
    ActiveSheet.Range("$A$1:" + c).AutoFilter Field:=19, Criteria1:="<>PEM" _
        , Operator:=xlAnd, Criteria2:="<>#N/A"
        
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Amount LC"
    
End Sub
    
Sub GR_pivot()
Attribute GR_pivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GR_pivot Macro
'

'
    Dim max_row As String
    Dim max_col As String
    max_row = CStr(getMaxRow(1))
    max_col = CStr(getMaxCol(1))
    
 '   Range("A1:T" + max_row).Select
    Range(Cells(1, 1), Cells(CDbl(max_row), CDbl(max_col))).Select
    
    Sheets.Add.Name = "PivotTable"
    
    Dim fileName As String
    fileName = ActiveWorkbook.Name
    Dim source As String
    source = "Sheet1!R1C1:R" + max_row + "C" + max_col
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6).CreatePivotTable _
        TableDestination:="PivotTable!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=6
    Sheets("PivotTable").Select
    
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Buyer")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Supplier")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Amount LC"), "Sum of Amount LC", xlSum
        
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/1000000"
    
    Range("C4").Select
    Selection.NumberFormat = "0.000"
    Selection.AutoFill Destination:=Range("C4:C" + CStr(getMaxRow(1)))
    
    Range("A4").Select
    ActiveWindow.FreezePanes = True
    With ActiveWindow
        .SplitRow = 3
    End With
    ActiveWindow.FreezePanes = True
    
End Sub

Sub OH_template()
'
' OH_template Macro
'

'
    Dim max_row As String
    Dim max_col As String
    max_row = getMaxRow(1)
    max_col = getMaxCol(1)
    Cells(1, CInt(max_col + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
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
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "Ok"
    Cells(2, CInt(max_col + 1)).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-8]+RC[-6]"
    Cells(2, CInt(max_col + 1)).Select
    
    Selection.AutoFill Destination:=Range(Cells(2, CInt(max_col + 1)), Cells(max_row, max_col + 1))
    Cells(1, CInt(max_col + 2)).Select
    ActiveCell.FormulaR1C1 = "SLoc"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
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
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Cells(2, CInt(max_col + 2)).Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(RC[-11],1)=""L"",""Prod"",IF(RC[-11]=""0012"",""LTB"",IF(LEFT(RC[-11],1)=""9"",""Quarantine"",IF(RC[-11]="""",""SubCon"",""WH""))))"
 '   =IF(LEFT(B2;1)="L";"Prod";IF(B2="0012";"LTB";IF(LEFT(B2;1)="9";"Quarantine";"WH"))) ;IF(B2="";"";"WH")
    Cells(2, CInt(max_col + 2)).Select
    Selection.AutoFill Destination:=Range(Cells(2, CInt(max_col + 2)), Cells(max_row, CInt(max_col + 2)))
    Columns(findCellInColumn(1, "Material")).Select
    Selection.NumberFormat = "@"
End Sub

Sub OH_pivot()
'
' OH_pivot Macro
'
    Dim max_row As String
    Dim max_col As String
    max_row = getMaxRow(1)
    max_col = getMaxCol(1)
    Range(Cells(1, 1), Cells(CDbl(max_row), CDbl(max_col))).Select
    Dim fileName As String
    fileName = ActiveWorkbook.Name
    
    Dim source As String
    source = "Sheet1!R1C1:R" + CStr(max_row) + "C" + CStr(max_col)
    Sheets.Add.Name = "PivotTable"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6).CreatePivotTable TableDestination:= _
        "PivotTable!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("PivotTable").Select
    
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SLoc")
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("WH").Position = 1
        .PivotItems("Prod").Position = 2
        .PivotItems("LTB").Position = 3
        .PivotItems("Quarantine").Position = 4
        .PivotItems("SubCon").Position = 5
    End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Ok"), "Sum of Ok", xlSum
        
    ActiveSheet.PivotTables("PivotTable1").RowAxisLayout xlTabularRow
    ActiveWorkbook.ShowPivotTableFieldList = False
 
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "Blocked"
  
    Range("H3:H4").Select
    Selection.Merge
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(Sheet1!C[-7],PivotTable!RC[-7],Sheet1!C)"
    Selection.AutoFill Destination:=Range("H5:H" + CStr(getMaxRow(1)))
    
    Range("H" + CStr(getMaxRow(1))).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-" + CStr(getMaxRow(1) - 5) + "]C:R[-1]C)"
    
    Range("A5").Select
    ActiveWindow.FreezePanes = True
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 4
    End With
    ActiveWindow.FreezePanes = True
    Columns("B:H").Select
    Selection.Style = "Comma [0]"
End Sub

Sub BreakLinks()
    Dim Links As Variant
    Links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    For i = 1 To UBound(Links)
    ActiveWorkbook.BreakLink _
        Name:=Links(i), _
        Type:=xlLinkTypeExcelLinks
    Next i
End Sub

Sub RemoveFormula()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Sheet1")

    With ws.UsedRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With
End Sub
