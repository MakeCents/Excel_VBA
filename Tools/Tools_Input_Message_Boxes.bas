Attribute VB_Name = "Tools_Input_Message_Boxes"
'The first two are from Microsoft. http://support.microsoft.com/kb/142141
'The following ones are modifications I use.
'=============================================================================================
Sub Using_InputBox_Method()
    'Using input and response
      Dim Response As Integer

      ' Run the Input Box.
      Response = Application.InputBox("Enter a number.", _
         "Number Entry", , 250, 75, "", , 1)

      ' Check to see if Cancel was pressed.
      If Response <> False Then

         ' If not, write the number to the first cell in the first sheet.
         Worksheets(1).Range("A1").Value = Response

      End If

End Sub
'=============================================================================================
Sub Using_InputBox_Function()
    'Using an input function, more like you were talking about.
      Dim Show_Box As Boolean
      Dim Response As Variant

      ' Set the Show_Dialog variable to True.
      Show_Box = True

      ' Begin While loop.
      While Show_Box = True

         ' Show the input box.
         Response = InputBox("Enter a number.", _
            "Number Entry", , 250, 75)

         ' See if Cancel was pressed.
         If Response = "" Then

            ' If Cancel was pressed,
            ' break out of the loop.
            Show_Box = False
         Else
            ' Test Entry to find out if it is numeric.
            If IsNumeric(Response) = True Then
               ' Write the number to the first
               ' cell in the first sheet in the active
               ' workbook.
               Worksheets(1).Range("a1").Value = Response
               Show_Box = False
            Else
               ' If the entry was wrong, show an error message.
               MsgBox "Please Enter Numbers Only"
            End If
         End If
      ' End the While loop.
      Wend
End Sub
'=============================================================================================
'=============================================================================================
Sub Using_InputBox_2()
    'Using a message box with simple input.
    MessageJ = "Question or statement?." ' Set prompt.
    TitleJ = "Box Title" ' Set title.
    DefaultJ = "What shows up in textbox initially." ' Set default.
    
    ' Display message, title, and default value.
    Answer = InputBox(MessageJ, TitleJ, DefaultJ)
    
    If Answer = "" Then
        'Do something if the text equals something.
    End If
    If StrPtr(Answer) = 0 Then
        'Do something if cancel is hit.
    Else
        'Do this after Okay is hit
    End If
End Sub
Sub Using_InputBox_3()
    'Using a message box with simple input
    'Display message
    Answer = InputBox("Question or statement?.", "Box Title", "What shows up in textbox initially.")
    
    If Answer = "" Then
        'Do something if the text equals something.
    End If
    If StrPtr(Answer) = 0 Then
        'Do something if cancel is hit.
    Else
        'Do this after Okay is hit
    End If
End Sub
'=============================================================================================
Sub MessageBox_vbyesnocancel()
    'Using msgboxes for yes/no answers.
    'the vbYesNoCancel can be changed to many things. Go to it and click Ctrl+J.
    Answer = MsgBox("Is the answer yes or no?", vbYesNoCancel)
        'Do this if it is no
    If Answer = vbNo Then
        'Do this for No
    End If
    If Answer = vbYes Then
        'Do this for Yes
    End If
    If Answer = vbCancel Then
        'Do this for Cancel
    End If
    
End Sub
