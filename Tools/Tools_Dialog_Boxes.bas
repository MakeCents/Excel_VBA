Attribute VB_Name = "Tools_Dialog_Boxes"
Sub Folder_Picker()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Show
        If .SelectedItems.Count = 0 Then GoTo 1
        fdlr = .SelectedItems(1)
1   End With
End Sub
Sub File_Picker()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Text", "*.txt", 1
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Show
        If .SelectedItems.Count = 0 Then GoTo 1
        fldr = .SelectedItems(1)
1   End With
End Sub

Sub SaveAs()
    With Application.FileDialog(msoFileDialogSaveAs)
        .Show
        .InitialFileName = ActiveWorkbook.Path & "\"
        If .SelectedItems.Count = 0 Then GoTo 1
        fldr = .SelectedItems(1)
1   End With
End Sub
Sub ChooseFile()
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'the number of the button chosen
Dim FileChosen As Integer
FileChosen = fd.Show
fd.title = "Select File"
fd.InitialFileName = ActiveWorkbook.Path
fd.InitialView = msoFileDialogViewSmallIcons
fd.Filters.Clear
fd.Filters.Add "Text", "*.txt"
fd.FilterIndex = 1
fd.ButtonName = "&MakeCents"
If FileChosen <> -1 Then
    'What to do if they hit cancel
Else
    fldr = fd.SelectedItems(1)
End If
End Sub
