Attribute VB_Name = "BrowseButton"
Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
Sub SelectFile()
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
    dialogBox.InitialFileName = Range("A12").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
    
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("B2").Value = dialogBox.SelectedItems(1)
        
        Select Case UCase(Left(GetFilenameFromPath(dialogBox.SelectedItems(1)), 2))
            Case "OH"
                ActiveSheet.OptionButtons("Option Button 13").Value = True
            Case "JI"
                ActiveSheet.OptionButtons("Option Button 7").Value = True
            Case "BA"
                ActiveSheet.OptionButtons("Option Button 8").Value = True
            Case "FG"
                ActiveSheet.OptionButtons("Option Button 9").Value = True
            Case "DA"
                ActiveSheet.OptionButtons("Option Button 10").Value = True
            Case "GR"
                Select Case Left(GetFilenameFromPath(dialogBox.SelectedItems(1)), 4)
                    Case "GR 1"
                        ActiveSheet.OptionButtons("Option Button 11").Value = True
                    Case "GR 4"
                        ActiveSheet.OptionButtons("Option Button 12").Value = True
                End Select
        End Select
    End If
    
End Sub
