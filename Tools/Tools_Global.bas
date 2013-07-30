Attribute VB_Name = "Tools_Global"

Sub Formulas_Off()
    Application.Calculation = xlManual
End Sub
Sub Formulas_On()
    Application.Calculation = xlAutomatic
End Sub
Sub ScreenUpdating_off()
    Application.ScreenUpdating = False
End Sub
Sub ScreenUpdating_on()
    Application.ScreenUpdating = True
End Sub
