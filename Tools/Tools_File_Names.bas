Attribute VB_Name = "Tools_File_Names"
Sub Copyfile()
    'copy and rename file in process
    file = "C:\Users\dgillespie\Desktop\New folder\Test2.xlsx"
    nfile = "C:\Users\dgillespie\Desktop\New folder\copy here\Test3.xlsx"
    FileCopy file, nfile
End Sub

'=======================================================================
'This will make figure folders, as many as desired.
'=======================================================================
Sub Make_Folders()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Show
        If .SelectedItems.Count = 0 Then GoTo 1
        fdlr = .SelectedItems(1)
1   End With
    'Using a message box with simple input
    'Display message
    Answer = InputBox("How many folders?", "Folders?", "5")
    If Answer = "" Then
        'Do something if the text equals something.
    End If
    If StrPtr(Answer) = 0 Then
        'Do something if cancel is hit.
    Else
        'Do this after Okay is hit
        For i = 1 To Answer
            Write_Folders (i), (fdlr)
        Next i
    End If
End Sub
Sub Write_Folders(i As Integer, fdlr As Variant)
            On Error Resume Next
            MkDir fdlr & "\" & "Figure " & i
End Sub
'=======================================================================
'This will make figure folders, as many as desired.
'=======================================================================

'=======================================================================
'This will rename a file
'=======================================================================
Sub Rename_File()
    On Error GoTo 1
    'Set names and then name.
    'This requires the entire path, which means it can move files during a rename.
    'This requires the file extension, which means you can change the file type during a rename.
        'Changing file type does not convert the file.
    
    'fPath is in A1
    fPath = Cells(1, 1)
    'fPath2 is if you want to move it, but for now it stays the same as fPath
    If Cells(1, 2) = "" Then
        fpath2 = fPath
    Else
        If Right(cells1, 2) = "\" Then
            fpath2 = Cells(1, 2)
        Else
            fpath2 = Cells(1, 2) & "\"
        End If
    End If
    'Start on the second row
    i = 2
    'Continue until the cell is empty
    Do While Not IsEmpty(Cells(i, 1))
        'List of files in column A
        oldfilename = fPath & Cells(i, 1).Value
        'List of what you want them named in column B
        newfilename = fpath2 & Cells(i, 2).Value
        'Name the old files the new file names
        Name oldfilename As newfilename
    'Add 1 to i to go to the next row
    i = i + 1
    'Do it again, if the cell is not empty
    Loop
1
End Sub
'=======================================================================
'=======================================================================

'=======================================================================
'This imports the files of type specified from a folder and subfolders.
'This starts and continues to the next double lines.
'=======================================================================
Sub Import_subfiles()
'
'This calls
'
    'Keeps Excel from freezing
    Application.Calculation = xlManual
    'Sets sheet name
    ws = ActiveSheet.Name
    Worksheets(ws).Select
    'Makes sure you want to clear this sheet
    If MsgBox("This will clear " & ws & ". Would you like to continue?", vbYesNo) = vbNo Then GoTo 2
    Range("A1:AA65000").Select
    Selection.ClearContents
    'Asks for the file type you want to import
    Answer = InputBox("What file type?" & _
    Chr(10) & "* = All" & _
    Chr(10) & "xls" & _
    Chr(10) & "doc" & _
    Chr(10) & "sgm" & _
    Chr(10) & "xlsx" & _
    Chr(10) & "txt", "File Type?", "*")
    If StrPtr(Answer) = 0 Then
        Application.Calculation = xlAutomatic
        GoTo 1
    End If
    'Asks for main folder to start in
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then GoTo 1
        fldr = .SelectedItems(1)
        Range("A1").Value = fldr & "\"
        
    End With
    
Dim FileNameWithPath As Variant
Dim ListOfFilenamesWithPath As New Collection    ' create a collection of filenames

' Filling a collection of filenames (search Excel files including subdirectories)
Call Subfolder_Search(ListOfFilenamesWithPath, "" & fldr, "*." & Answer, True)

' Print list to immediate debug window and as a message window
    i = 2
For Each FileNameWithPath In ListOfFilenamesWithPath    ' cycle for list(collection) processing
        Debug.Print FileNameWithPath & Chr(13)
        'Removes the path from the file name
        NFileNameWithPath = Mid(FileNameWithPath, Len(fldr) + 2, 1000)
        'Delete the N before "FileNameWithPath" in order to keep the whole path
        Worksheets(ws).Range("A" & i).Value = NFileNameWithPath
        i = i + 1
Next FileNameWithPath

' Print to immediate debug window and message if no file was found
If ListOfFilenamesWithPath.Count = 0 Then
    Debug.Print "No file was found !"
    MsgBox "No " & Answer & " was found!"
End If
  
1   Range("A1").Select
2   Application.Calculation = xlAutomatic


End Sub
Private Sub Subfolder_Search(dfoundfiles As Collection, dPath As String, dMask As String, dIncludeSubdirectories As Boolean)
'
'This is called
'
    'When there is a bad file name an error occurs, seen with ??????? in the name on the worksheet.
    On Error Resume Next
Dim DirFile As String
Dim CollectionItem As Variant
Dim SubDirCollection As New Collection

' Add backslash at the end of path if not present
dPath = Trim(dPath)
If Right(dPath, 1) <> "\" Then dPath = dPath & "\"

' Searching files accordant with mask
DirFile = Dir(dPath & dMask)
Do While DirFile <> ""
dfoundfiles.Add dPath & DirFile  'add file name to list(collection)
DirFile = Dir ' next file
Loop

' Procedure exiting if searching in subdirectories isn't enabled
If Not dIncludeSubdirectories Then Exit Sub

' Searching for subdirectories in path
DirFile = Dir(dPath & "*", vbDirectory)
Do While DirFile <> ""
    ' Add subdirectory to local list(collection) of subdirectories in path
    If DirFile <> "." And DirFile <> ".." Then If ((GetAttr(dPath & DirFile) And vbDirectory) = 16) Then SubDirCollection.Add dPath & DirFile
    DirFile = Dir 'next file
Loop

' Subdirectories list(collection) processing
For Each CollectionItem In SubDirCollection
     Call Subfolder_Search(dfoundfiles, CStr(CollectionItem), dMask, dIncludeSubdirectories) ' Recursive procedure call
Next

Application.Calculation = xlAutomatic
    Range("A1").Select
End Sub
'=======================================================================
'This is the end of the Subfolder importing
'=======================================================================

'=======================================================================
'This imports the files of type specified from a single folder.
'=======================================================================
Sub Import_File_Names()
Dim fileList() As String
Dim fName As String
Dim fPath As String
Dim i As Integer
Dim startrow As Integer
Dim ws As Worksheet
Dim filetype  As String
wsn = ActiveSheet.Name
     'Makes sure you want to clear this sheet
    If MsgBox("This will clear " & wsn & ". Would you like to continue?", vbYesNo) = vbNo Then GoTo 3
    Range("A1:AA65000").Select
    Selection.ClearContents
    Application.Calculation = xlManual
    T = InputBox("What file type?" & _
    Chr(10) & "* = All" & _
    Chr(10) & "xls" & _
    Chr(10) & "doc" & _
    Chr(10) & "sgm" & _
    Chr(10) & "xlsx" & _
    Chr(10) & "txt", "File Type?", "*")
    If StrPtr(T) = 0 Then
        Sheets(wsn).Select
        Application.Calculation = xlAutomatic
        GoTo 3
    End If
    '=======================================================
    'Sets sheet name
    Sheets(wsn).Select
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub
            fPath = .SelectedItems(1) & "\"
        End With
1   filetype = T
    Set ws = Worksheets(wsn)
    ws.Select
    startrow = 1    'starting row for the data
    '========================================================
    Range("A1").Value = fPath
    fName = Dir(fPath & "*." & filetype)
    While fName <> ""
        i = i + 1
        ReDim Preserve fileList(1 To i)
        fileList(i) = fName
        fName = Dir()
    Wend
    If i = 0 Then
        MsgBox "No " & T & " found in " & fPath & "       ", vbExclamation
        Application.Calculation = xlAutomatic
        Exit Sub
    End If
    For i = 1 To UBound(fileList)
        ws.Range("A" & i + startrow).Value = fileList(i)
    Next
    MsgBox "Files Imported"
3   Application.Calculation = xlAutomatic
    Range("A1").Select
End Sub
'=======================================================================
'=======================================================================

'=======================================================================
'Adds or removes the 8 digit reference number, based on PLISN,
'=======================================================================
Sub Convert_PLISN_to_Number_FileNamer()
Dim fileList() As String
Dim fName As String
Dim fPath As String
Dim i As Integer
Dim filetype  As String
     'Makes sure you want to clear this sheet
    T = InputBox("What file type?" & _
    Chr(10) & "* = All" & _
    Chr(10) & "xls" & _
    Chr(10) & "doc" & _
    Chr(10) & "sgm" & _
    Chr(10) & "xlsx" & _
    Chr(10) & "txt", "File Type?", "*")
    If StrPtr(T) = 0 Then GoTo 4
   
    '=======================================================
        On Error GoTo 3
       
       fPath = Dialog(L)
       If fPath = "0" Then GoTo 4
1
    filetype = T
    '========================================================
    fName = Dir(fPath & "*." & filetype)
    While fName <> ""
        i = i + 1
        ReDim Preserve fileList(1 To i)
        fileList(i) = fName
        fName = Dir()
    Wend
    
    If MsgBox("Do you want to remove converted numbers?", vbYesNo) = vbNo Then GoTo 5
    If i = 0 Then
        MsgBox "No " & T & " found in " & fPath & "       ", vbExclamation
        Application.Calculation = xlAutomatic
        Exit Sub
    End If
    For i = 1 To UBound(fileList)
        N = Mid(fileList(i), 10, 100)
        Name fileList(i) As N
    Next
    GoTo 4
    
5
    If MsgBox("Do you want to convert PLISN numbers?", vbYesNo) = vbNo Then
        MsgBox "Nothing Done", vbExclamation
        GoTo 4
    End If
    If i = 0 Then
        MsgBox "No " & T & " found in " & fPath & "       ", vbExclamation
        Application.Calculation = xlAutomatic
        Exit Sub
    End If
    For i = 1 To UBound(fileList)
        N = ""
        PLISN = Left(fileList(i), 4)
        Dim p As Variant
        For PTO = 1 To 4
            p = Mid(PLISN, PTO, 1)
            p1 = LettertoNumber(p)
            N = N & p1
        Next PTO
        Name fileList(i) As N & "_" & fileList(i)
    Next
    MsgBox "Done"
    GoTo 4
3   MsgBox "Did not remove all of the converted numbers!", vbCritical
4

End Sub
Function LettertoNumber(p As Variant)
    For i = 48 To 65
    If p = Chr(i) Then LettertoNumber = i - 12
    Next i
    For i = 65 To 130
    If p = Chr(i) Then LettertoNumber = i - 55
    Next i
1
End Function
Function Dialog(L As Variant)
    On Error GoTo 1
     With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            If .SelectedItems.Count = 0 Then GoTo 2
            fPath = .SelectedItems(1) & "\"
        End With
    Dialog = fPath
    GoTo 2
1   Dialog = 0
2
End Function
'=======================================================================
'End of add or remove 8 digit reference number
'=======================================================================

'=======================================================================
'Adds or removes the 8 digit reference number, based on PLISN, a folder
    'and subfolders.
'This starts and continues to the next double lines.
'=======================================================================
Sub Convert_PLISN_to_Number_FileNamer_Subfiles()
'
'This calls
'
    On Error GoTo 1
    'Asks for the file type you want to import
    Answer = InputBox("What file type?" & _
    Chr(10) & "* = All" & _
    Chr(10) & "xls" & _
    Chr(10) & "doc" & _
    Chr(10) & "sgm" & _
    Chr(10) & "xlsx" & _
    Chr(10) & "txt", "File Type?", "*")
    If StrPtr(Answer) = 0 Then GoTo 2

    'Asks for main folder to start in
    
    fldr = Dialog2(L)
    If fldr = "0" Then GoTo 2

    
Dim FileNameWithPath As Variant
Dim ListOfFilenamesWithPath As New Collection    ' create a collection of filenames

' Filling a collection of filenames (search Excel files including subdirectories)
Call FSubfolder_Search(ListOfFilenamesWithPath, "" & fldr, "*." & Answer, True)

' Print list to immediate debug window and as a message window
    i = 2
   
    If MsgBox("Do you want to remove the numbers?", vbYesNo) = vbYes Then
        V = "A"
        GoTo 9
    End If
    If MsgBox("Do you want to add numbers?", vbYesNo) = vbYes Then
        V = "B"
        GoTo 9
    End If
    GoTo 1
9
For Each FileNameWithPath In ListOfFilenamesWithPath    ' cycle for list(collection) processing
        Debug.Print FileNameWithPath & Chr(13)
        'Removes the path from the file name
        NFileNameWithPath = Mid(FileNameWithPath, Len(fldr) + 2, 1000)
        'Delete the N before "FileNameWithPath" in order to keep the whole path
        
        'Do this to add numbers
        
6
    For b = 1 To 1000
        If Mid(FileNameWithPath, Len(FileNameWithPath) - b, 1) = "\" Then
            FNName = Mid(FileNameWithPath, Len(FileNameWithPath) - b + 1, 1000)
            FNfLdr = Mid(FileNameWithPath, 1, Len(FileNameWithPath) - b)
            GoTo 7
        End If
    Next b
    
7   If V = "A" Then
        FNName = Mid(FNName, 10, 1000)
        GoTo 8
    End If
        NName = ""
        PLISN = Left(FNName, 4)
        Dim p As Variant
        For PTO = 1 To 4
            p = Mid(PLISN, PTO, 1)
            p1 = LettertoNumber2(p)
            NName = NName & p1
        Next PTO
        NName = NName & "_"
8
    Name FileNameWithPath As FNfLdr & "\" & NName & FNName

Next FileNameWithPath

' Print to immediate debug window and message if no file was found
If ListOfFilenamesWithPath.Count = 0 Then
    Debug.Print "No file was found !"
    MsgBox "No " & Answer & " was found!"
End If
GoTo 2
1   MsgBox "Error", vbCritical
2

End Sub
Function LettertoNumber2(p As Variant)
    For i = 48 To 65
    If p = Chr(i) Then LettertoNumber2 = i - 12
    Next i
    For i = 65 To 130
    If p = Chr(i) Then LettertoNumber2 = i - 55
    Next i
1
End Function
Function Dialog2(L As Variant)
    On Error GoTo 1
     With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            If .SelectedItems.Count = 0 Then GoTo 2
            fPath = .SelectedItems(1) & "\"
        End With
    Dialog2 = fPath
    GoTo 2
1   Dialog2 = 0
2
End Function
Private Sub FSubfolder_Search(dfoundfiles As Collection, dPath As String, dMask As String, dIncludeSubdirectories As Boolean)
'
'This is called
'
    'When there is a bad file name an error occurs, seen with ??????? in the name on the worksheet.
    On Error Resume Next
Dim DirFile As String
Dim CollectionItem As Variant
Dim SubDirCollection As New Collection

' Add backslash at the end of path if not present
dPath = Trim(dPath)
If Right(dPath, 1) <> "\" Then dPath = dPath & "\"

' Searching files accordant with mask
DirFile = Dir(dPath & dMask)
Do While DirFile <> ""
dfoundfiles.Add dPath & DirFile  'add file name to list(collection)
DirFile = Dir ' next file
Loop

' Procedure exiting if searching in subdirectories isn't enabled
If Not dIncludeSubdirectories Then Exit Sub

' Searching for subdirectories in path
DirFile = Dir(dPath & "*", vbDirectory)
Do While DirFile <> ""
    ' Add subdirectory to local list(collection) of subdirectories in path
    If DirFile <> "." And DirFile <> ".." Then If ((GetAttr(dPath & DirFile) And vbDirectory) = 16) Then SubDirCollection.Add dPath & DirFile
    DirFile = Dir 'next file
Loop

' Subdirectories list(collection) processing
For Each CollectionItem In SubDirCollection
     Call FSubfolder_Search(dfoundfiles, CStr(CollectionItem), dMask, dIncludeSubdirectories) ' Recursive procedure call
Next


End Sub
'=======================================================================
'End of add or remove 8 digit reference number from Subfolders
'=======================================================================

