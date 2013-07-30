Attribute VB_Name = "Tools_Text_Editing"
'This section is devoted to finding and seperating text based on a repeated character.
'Also known as slicing
'=====================================================================================
Sub Find_last_of_something()
'Place in column that follows information to find
    ActiveCell.FormulaR1C1 = _
        "=FIND(""*"",SUBSTITUTE(RC[-1],""\"",""*"",LEN(RC[-1])-LEN(SUBSTITUTE(RC[-1],""\"",""""))))"
End Sub
Sub Find_last_and_After()
'Place in 2nd column away from information to find
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[-2],FIND(""*"",SUBSTITUTE(RC[-2],""\"",""*"",LEN(RC[-2])-LEN(SUBSTITUTE(RC[-2],""\"",""""))),1)+1,1000)"
End Sub
Sub Find_last_and_before()
'Place in column that follows information to find
    ActiveCell.FormulaR1C1 = _
        "=LEFT(RC[-1],FIND(""*"",SUBSTITUTE(RC[-1],""\"",""*"",LEN(RC[-1])-LEN(SUBSTITUTE(RC[-1],""\"",""""))),1))"
End Sub
'=====================================================================================
'Replace ", " with "," and add " " * # of ", " found.
'=====================================================================================
Sub Replace_comma_spaces()
    Application.Calculation = xlManual
    c = 12
    For c = 12 To 115
    i = 3
    Do While Not IsEmpty(Cells(i, c))
        T = Cells(i, c)
        f = Replace(Cells(i, c), ", ", ",", 1, , vbTextCompare)
        m = Len(T) - Len(f)
        Cells(i, c) = f & Space(m)
    i = i + 1
    Loop
    If c = 115 Then
        Application.Calculation = xlAutomatic
        Exit Sub
    End If
    c = 114
    Next
End Sub
'=====================================================================================

