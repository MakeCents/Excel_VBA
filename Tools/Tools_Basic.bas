Attribute VB_Name = "Tools_Basic"
Sub Navigation_Tools()
'Tools used to navigate worksheet cells

    'Selects specific cell
    Range("A1").Select
    
    'Selects specific range
    Range("A1:B3").Select
    
    'Within your selection will be an active cell, which can be adjusted without loosing your selection
    Range("B2").Activate
    
    'Moves Row +1 and Column +1 from activecell
    ActiveCell.Offset(1, 1).Select
    
    'Selects from activecell to last cell with something in it, in a direction
        'Change "Down" to, "ToRight", "ToLeft", or "Up".
    Range(Selection, Selection.End(xlDown)).Select
    
    'Goes to last cell that is not empty
        'Change "Down" to, "ToRight", "ToLeft", or "Up".
    Selection.End(xlDown).Select

    'Takes selection you have and adds 4 more columns, from left column, to selection
        'If more than four columns selected then it doesn't do anything
        'Modify ActiveCell.row or ActiveCell.Column + 4 for other results
    Range(Selection, Cells(ActiveCell.Row, ActiveCell.Column + 4)).Select
    
    'Navigates to the last (Row, Column) of the worksheet that has had information in it
        'I don't like this one
    ActiveCell.SpecialCells(xlLastCell).Select
    
    'Navigates sheets
        'Can replace Next with Previous
    ActiveSheet.Next.Select
    
End Sub
Sub Search_Tools()

'Find tools

    'Errors are recieved when not found. Make sure to check search variables.
    On Error Resume Next
        '"X" can be what you want to find.
        'Values can be formulas or comments
        'False can be true (MatchCase, SearchFormat)
        'Columns can be Rows
        'After:=ActiveCell can be specific
        'Searchdirection can be previous
        'Whole can be part
    Cells.Find(What:="X", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    
    'Find has been established, find next
    Cells.FindNext(After:=ActiveCell).Activate

'Replace tools

    'Will not return an error if the activecell does not have what it is looking for
    'Replaces the current "X" with "Y"
    'This is more if you want to replace something within a cell.
    'Requires a Find tool first or replace "ActivceCell" with something more specific, Range("A1")
    ActiveCell.Replace What:="X", Replacement:="Y", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Errors for not finding things are not recieved during a replace all
    'Replaces all "X" with "Y"
    Cells.Replace What:="X", Replacement:="Y", LookAt:=xlPart, SearchOrder _
        :=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
