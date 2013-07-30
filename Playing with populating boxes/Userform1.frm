VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Box Population"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   OleObjectBlob   =   "Userform1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    
    Dim NC As New Collection
    ' Start at first row
    i = 1
    k = Cells(i, TextBox1)
    ' Do until blank cell in first column
    Do While Not IsEmpty(Cells(i, TextBox1))
        ' Add Cell values to collection
        NC.Add Cells(i, TextBox1)
        ' Each item in first column added to collection will now be added to combobox1
        With ComboBox1
            .AddItem Cells(i, TextBox1)
        End With
        With ListBox1
            .AddItem Cells(i, TextBox1)
        End With
        
        ' Next cell
        ComboBox1.ListRows = ComboBox1.ListRows + i
        If ComboBox1.ListRows > 40 Then ComboBox1.ListRows = 40
        i = i + 1
    Loop
    l = Cells(i - 1, TextBox1)
    
1
        
    
    'This shows the whole list.
    MsgBox k & " thru " & l & " Added", vbExclamation
End Sub

Private Sub CommandButton2_Click()
    k = 0
    l = 0
    i = 1
    Do While Not IsEmpty(Cells(i, 1))
        P = Cells(i, 1)
        m = m & Chr(10) & P
        k = k + 1
        If k = l + 5 Then
            MsgBox m, vbInformation
            l = k
            m = ""
        End If
        i = i + 1
    Loop
    
    
End Sub

Private Sub CommandButton3_Click()
    ComboBox1.Clear
    ListBox1.Clear
    Dim NC As New Collection
    ' Start at first row
    i = 1
    ' Do until blank cell in first column
    Do While Not IsEmpty(Cells(i, 1))
        ' Add Cell values to collection
        NC.Add Cells(i, 1)
        ' Each item in first column added to collection will now be added to combobox1
        With ComboBox1
            .AddItem Cells(i, 1)
        End With
        With ListBox1
            .AddItem Cells(i, 1)
        End With
        ComboBox1.ListRows = i
        i = i + 1
    Loop
End Sub

Private Sub CommandButton4_Click()
    Dim NC As New Collection
    k = 1
    l = 0
    i = 1
    Do While Not IsEmpty(Cells(i, 1))
        
       ' Add Cell values to collection
        NC.Add Cells(i, 1)
        ' Each item in first column added to collection will now be added to combobox1
        With ComboBox1
            .AddItem Cells(i, 1)
        End With
        With ListBox1
            .AddItem Cells(i, 1)
        End With
        
        i = i + 1
    Loop
    For Each CollectionItem In NC
        If m = "" Then
            m = CollectionItem
            GoTo 1
        End If
        m = m & Chr(10) & CollectionItem
1       If k = l + 5 Then
            MsgBox m, vbInformation
            l = k
        End If
        k = k + 1
        
    Next
End Sub



Private Sub UserForm_Initialize()
    ComboBox1.Clear
    ListBox1.Clear
    Dim NC As New Collection
    ' Start at first row
    k = 0
    l = 0
    i = 1
    ' Do until blank cell in first column
    Do While Not IsEmpty(Cells(i, 1))
        ' Add Cell values to collection
        NC.Add Cells(i, 1)
        ' Each item in first column added to collection will now be added to combobox1
        With ComboBox1
            .AddItem Cells(i, 1)
        End With
        With ListBox1
            .AddItem Cells(i, 1)
        End With
        
        ' This will message box each increment of 5
        P = Cells(i, 1)
        m = m & Chr(10) & P
        k = k + 1
        If k = l + 5 Then
            MsgBox m, vbInformation
            l = k
            m = ""
        End If
       
        ' Next cell
        ComboBox1.ListRows = i
        i = i + 1
    Loop
    m = ""
    l = 0
    k = 1
    ' The following would be done if there was not a place to put this information yet.
    ' This will message box a consolodated list of each set of 5
    For Each CollectionItem In NC
        If m = "" Then
            m = CollectionItem
            GoTo 1
        End If
        m = m & Chr(10) & CollectionItem
1       If k = l + 5 Then
            MsgBox m, vbInformation
            l = k
        End If
        k = k + 1
        
    Next
    'This shows the whole list.
    MsgBox m, vbInformation
    MsgBox "First column populated.", vbExclamation
End Sub
