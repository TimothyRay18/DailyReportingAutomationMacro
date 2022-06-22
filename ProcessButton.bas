Attribute VB_Name = "ProcessButton"
Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Sub Process()
    Dim r_button As String
    r_button = Range("B4").Value

    Dim dir As String
    dir = Range("B2").Value
    
    CleanWorkbook

    Dim wb As Workbook
    
    Dim opened_wb As String
    
    fn = GetFilenameFromPath(dir)
    
    Select Case r_button
    Case 7
   '     MsgBox "OH"
        response = MsgBox("On Hands using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Set wb = Workbooks.Open(dir)
        wb.Worksheets("Sheet1").Activate
        
        opened_wb = ActiveWorkbook.Name
        
        CopyData 1, 1, ThisWorkbook.Name
        
        Controller.OH_template
        Controller.OH_pivot
    Case 1
        response = MsgBox("JIT using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Workbooks.OpenText fileName:= _
        dir, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 2), Array(5, 1), Array(6, 2), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 2), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), _
        Array(49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array( _
        55, 1), Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), _
        Array(62, 1), Array(63, 1), Array(64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array( _
        68, 1), Array(69, 1), Array(70, 1), Array(71, 1), Array(72, 1), Array(73, 1)), _
        TrailingMinusNumbers:=True
        opened_wb = ActiveWorkbook.Name
        
        CopyData 9, 4, ThisWorkbook.Name
        
        Controller.JIT
    Case 2
     '   MsgBox "BF"
     
        response = MsgBox("Backflush using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Workbooks.OpenText fileName:= _
        dir, Origin:= _
        xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 2), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:=True
        opened_wb = ActiveWorkbook.Name
        
        CopyData 4, 2, ThisWorkbook.Name
        
        Controller.BF
    Case 3
     '   MsgBox "FG"
     
        response = MsgBox("Finish Good using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Workbooks.OpenText fileName:= _
        dir, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 2), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1)), TrailingMinusNumbers:=True
        opened_wb = ActiveWorkbook.Name
        
        CopyData 16, 2, ThisWorkbook.Name
        
        Controller.FG
    Case 4
     '   MsgBox "DI"
        
        response = MsgBox("Daily Inventory using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Workbooks.OpenText fileName:= _
        dir, Origin:= _
        xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1)), TrailingMinusNumbers:=True
        opened_wb = ActiveWorkbook.Name
        
        CopyData 11, 2, ThisWorkbook.Name
        
        Controller.DI
    Case 5
     '   MsgBox "GR101"
     
        response = MsgBox("GR 101 using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Workbooks.OpenText fileName:= _
        dir, Origin:=xlWindows _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 4), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1)), TrailingMinusNumbers:=True
        opened_wb = ActiveWorkbook.Name
        
        CopyData 4, 2, ThisWorkbook.Name
        
        Controller.GR_template (r_button)
        Controller.GR_pivot
        
    Case 6
    '    MsgBox "GR411"
        response = MsgBox("GR 411 using " + fn + vbNewLine + "Continue?", vbYesNo, "Confirmation")
 
        If response = vbNo Then
            Exit Sub
        End If
        
        Workbooks.OpenText fileName:= _
        dir, Origin:=xlWindows _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 4), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1)), TrailingMinusNumbers:=True
        opened_wb = ActiveWorkbook.Name
        
        CopyData 4, 2, ThisWorkbook.Name
        
        Controller.GR_template (r_button)
        Controller.GR_pivot
    End Select
    
    Application.DisplayAlerts = False
    Windows(opened_wb).Close
    Application.DisplayAlerts = True
End Sub

Sub CopyData(X As Double, Y As Double, file As String)
    lastrow = ActiveSheet.Cells(Rows.Count, Y).End(xlUp).row
    lastcol = ActiveSheet.Cells(X, Columns.Count).End(xlToLeft).Column
    
    ActiveSheet.Range("A1", Cells(lastrow + 1, lastcol + 1)).Select
    Selection.Copy
    
    Workbooks(file).Activate
    ActiveWorkbook.Worksheets("Sheet1").Paste

    ActiveWorkbook.Worksheets("Sheet1").Activate
End Sub

Sub CleanWorkbook()
    Dim Sh As Worksheet
    Application.DisplayAlerts = False
    For Each Sh In Worksheets
        If Sh.Name <> ActiveSheet.Name And Sh.Name <> "Master Data" Then Sh.Delete
    Next Sh
    Application.DisplayAlerts = True
    Sheets.Add.Name = "Sheet1"
End Sub

