Attribute VB_Name = "Tools_Write_File"
Sub Write_to_column()

'A single apostrophe, ', represents a note
'Replace all double apostrophe, '', represents code that can be used, with nothing to use this
'Relace all apostrophe, space, and apostrophe, , represents more options in code, with nothing to use this
'To change it to write row first then column, replace (i, t) with (t, i) and (i, t + 1) with (t + 1, i)
'The previous two used together will transpose data
'Syntax Errors are on purpose, please read notes when this happens.
'WARNING - Only replace '' or ' '
'=======================================================================================================
    On Error GoTo 1
    'Set varialbes via boxes
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Show
        fdlr = .SelectedItems(1)
    End With
    N = InputBox("What do you want to name the file?", "Save As", ActiveWorkbook.Name)
    If StrPtr(N) = 0 Then GoTo 1
    E = InputBox("What file extension?", "Extension", "txt")
    If StrPtr(E) = 0 Then GoTo 1
    file = fdlr & "\" & N & "." & E
    'Set variables manually
    'P = ActiveWorkbook.path & "\" 'Same path as workbook
    'N = "Test" 'Name of file
    'E = "." & "txt" 'Extension
    'file = P & N & E
    '===================================================================================================
    'Use "For Append As" in place of "For Output Access Write As" to add or append information to file
2   Open file For Output Access Write As #1 'Finds/Creates the file path + name + extention
    '===================================================================================================
    Worksheets("Sheet1").Select 'Select the sheet
    '===================================================================================================
    i = 1 'This is the row to strat with from the top
    T = 1 'This is the column to start in from left to right
    ' 'R = 0 'This counts the rows
    ' 'LC = 0 'This counts the columns
    '===================================================================================================
    'If the cell is empty then it will stop looping
    Do While Not IsEmpty(Cells(i, T))
        'Use '' to print all in row before going to next column
        'If the cell is empty then it will stop looping
        ''Do While Not IsEmpty(Cells(i, t))
            'Write the cells value in the file
            Print #1, Cells(i, T) ' 'Put a ' before this row. This row is not used for tabs and tables
            'Add 1 to go to the next column
            ''t = t + 1
        'Go to the second Do While
        ''Loop
    '===================================================================================================
        'Use  to print all in row before going to next column with tabs between cells for tables
        ' 'Do While Not IsEmpty(Cells(i, t))
            ' 'If Cells(i, t + 1) = Empty Then
                 'This keeps from adding a tab when there is not another cell.
                 ' 'L = L & Cells(i, t)
             ' 'Else
                 'This adds a tab for the next cell.
                 ' 'L = L & Cells(i, t) & vbTab
             ' 'End If
             ' 't = t + 1
         ' 'Loop
         ' 'c = t - 1
         ' 'Print #1, L
         ' 'L = ""
         ' 't = 1
         ' 'R = R + 1
         ' 'If c > LC Then LC = c
    '===================================================================================================
        'Reset the column
        ''t = 1
    'Add 1 to go to the next row
    i = i + 1
    'Go back to the 1st Do While
    Loop
    ' 'Print #1, "Rows = " & R & vbNewLine & "Columns = " & LC
    Close #1
1
End Sub
