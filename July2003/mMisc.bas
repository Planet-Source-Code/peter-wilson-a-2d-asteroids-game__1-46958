Attribute VB_Name = "mMisc"
Option Explicit

Private Const g_sngPIDivideBy180 = 0.0174533!

Public Sub DrawCrossHairs(CurrentForm As Form)

    CurrentForm.DrawStyle = vbSolid
    CurrentForm.DrawWidth = 1
    
    ' Draw vertical line.
    CurrentForm.ForeColor = RGB(0, 48, 0)
    CurrentForm.Line (CurrentForm.ScaleWidth / 2, 0)-(CurrentForm.ScaleWidth / 2, CurrentForm.ScaleHeight)

    ' Draw horizontal line.
    CurrentForm.ForeColor = RGB(0, 24, 0)
    CurrentForm.Line (0, CurrentForm.ScaleHeight / 2)-(CurrentForm.ScaleWidth, CurrentForm.ScaleHeight / 2)
    
End Sub
Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Public Sub Debug_PrintMatrix(m1 As mdrMATRIX3x3)

    With m1
        
        frmCanvas.ForeColor = RGB(192, 192, 192)
        frmCanvas.CurrentX = 0
        frmCanvas.CurrentY = 0
        
        frmCanvas.Print .rc11, .rc12, .rc13
        frmCanvas.Print .rc21, .rc22, .rc23
        frmCanvas.Print .rc31, .rc32, .rc33
        
    End With
    
End Sub

Public Function GetRNDNumberBetween(Min As Variant, Max As Variant) As Single

    GetRNDNumberBetween = (Rnd * (Max - Min)) + Min

End Function

