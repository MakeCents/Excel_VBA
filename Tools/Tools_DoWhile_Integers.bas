Attribute VB_Name = "Tools_DoWhile_Integers"
Sub DoWhile_Process()
    Do Until Application.CalculationState = xlDone
       DoEvents
    Loop
End Sub
Sub DoWhile_Isempty()
    'i and t represent coordinates for "Cells()". Cells(1,1) = A1
    i = 1
    T = 1
    'This is looking at the cells, not selecting them, this is faster
    'Will look in the cells A1, then B2, then C3 and so on, as long as it is NOT empty
        'Delete Not to do it as long as the cell is empty
    Do While Not IsEmpty(Cells(i, T))
        'This will increase the row by one
        i = i + 1
        'This will increase the row by one
        T = T + 1
    Loop
End Sub

Sub DoWhile_Value()
    T = 1
    Do While T < 11
        Cells(T, 1) = T
        T = T + 1
    Loop
End Sub

Sub For_Integers()
    Dim i As Integer
    Dim T As Integer
    'This establishes the current cell coordinates
    a = ActiveCell.Row
    b = ActiveCell.Column
    'This will incrementally add 0 to t until it gets to 1, performing i twice, once for 0 and then for 1
    'Change t to 0 to 0 and it behaves as if it isn't there at all.
    For T = 0 To 1
    'This will incrementally add 1 to i until it gets to 10
    For i = 1 To 50
        'The following will navigate using i from activecell
        Cells(a - 1 + i, b + T).Select
        'The following will toggle the cells as i or "", just for an example
        If ActiveCell = i Then
            ActiveCell = ""
        Else
            ActiveCell = i
        End If
    Next i
    Next T
    Cells(a, b).Select
End Sub
'=========================================================================
'The following all do the same thing, three different ways.
'=========================================================================
Sub For_DoWhile()
    'Set t
    T = 1
    For i = 1 To 4
    Do While T < 11
        Cells(i, T) = T
        T = T + 1
    Loop
    'Reset t
    T = 1
    Next i
End Sub
Sub DoWhile_For()
    'Set t
    T = 1
    Do While T < 11
    For i = 1 To 4
        Cells(i, T) = T
    Next i
    T = T + 1
    Loop
End Sub
Sub For_For()
    For T = 1 To 10
    For i = 1 To 4
        'Switch the t and the i and it will count down instead of to the right.
        'Add "11 - " before the t after the = sign and it will cout backwards
        Cells(i, T) = T
    Next i
    Next T
End Sub
'=========================================================================
