Attribute VB_Name = "Tools_Write_csv_file"
Sub Write_CSV_file()
    'this will write a csv file where you want infrom column A starting at row 5 and then add numbers
    '   for importing to Corel.
    Dim Ref As New Collection
    i = 5
    Do While Not (Cells(i, 1)) = ""
        Ref.Add Item:=Cells(i, 1)
        i = i + 1
    Loop
    m = ""
    For Each Item In Ref
        If Item.Row - 4 = Ref.Count Then
            m = m & Item
        Else
            m = m & Item & ","
        End If
    Next
    m = m & ","
    For Each Item In Ref
        If Item.Row - 4 = Ref.Count Then
            m = m & Item.Row - 5
        Else
            m = m & Item.Row - 5 & ","
        End If
    Next
    
    Answer = InputBox("What do you want to name the file?", "Name CSV File", "")
    
    If Answer = "" Then
        'Do something if the text equals something.
        MsgBox "File name can not be nothing", vbCritical
        Exit Sub
    End If
    If StrPtr(Answer) = 0 Then
        'Do something if cancel is hit.
        Exit Sub
    Else
        'Do this after Okay is hit
        fName = Answer
    End If
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path & "\"
        .title = "Where do you want to save the file?"
        .Show
        If .SelectedItems.Count = 0 Then GoTo 1
        fdlr = .SelectedItems(1)
1   End With
    
    Open fName & ".csv" For Output Access Write As #1  '***** Change Directory name and file name to suit
        Print #1, m
    Close #1
End Sub
