Attribute VB_Name = "mRasterizer"
Option Explicit

Public Sub mdrLine(Xstart As Single, Ystart As Single, Xend As Single, Yend As Single)

    ' Bresenham's classic incremental line scan-conversion algorithm (slightly modified)
    ' ==================================================================================
    ' (I created this because I thought I could use it to implement a quick-n-nasty hidden-line
    '  removal technique for the Asteroids... however my original assumption was wrong.
    '  However, this code works perfectly so I've left it in for your viewing pleasure.)
    '
    ' Use this routine to draw lines between two points on the screen.
    
    Dim sngDeltaX As Single
    Dim sngDeltaY As Single
    
    Dim sngSlope As Single
    
    Dim sngXIncr As Single
    Dim sngYIncr As Single
    
    Dim sngOriginalX As Single
    Dim sngOriginalY As Single
    
    sngOriginalX = Xend
    sngOriginalY = Yend
    
    If (Xstart = Xend) And (Ystart = Yend) Then
        frmCanvas.PSet (Xend, Yend)
        Exit Sub
    End If
    
    sngDeltaX = (Xend - Xstart)
    sngDeltaY = (Yend - Ystart)
    
    If (sngDeltaX = 0) Or (sngDeltaY = 0) Then
        sngSlope = 0
    Else
        sngSlope = (sngDeltaY / sngDeltaX)
    End If
    
        
    If (Abs(sngSlope) < 1 And (Xstart > Xend)) Or _
       (Abs(sngSlope) > 1 And (Ystart > Yend)) Or _
       ((sngSlope = 0) And (Ystart > Yend)) Then
       
        Call SwapEndPoints(Xstart, Ystart, Xend, Yend)
    End If
    
    
    If (Abs(sngSlope) < 1) And (sngDeltaX <> 0) Then
        
        sngSlope = sngSlope * Screen.TwipsPerPixelX
        sngYIncr = Ystart
        For sngXIncr = Xstart To Xend Step Screen.TwipsPerPixelX
            frmCanvas.Refresh ' Good place to slow everything down and enjoy the pretty picture getting made.
            frmCanvas.PSet (sngXIncr, sngYIncr)
            sngYIncr = sngYIncr + sngSlope
        Next sngXIncr
        
    Else
    
        If sngDeltaX <> 0 Then sngSlope = 1 / sngSlope
        
        sngSlope = sngSlope * Screen.TwipsPerPixelY
        sngXIncr = Xstart
        For sngYIncr = Ystart To Yend Step Screen.TwipsPerPixelY
            frmCanvas.Refresh ' Good place to slow everything down and enjoy the pretty picture getting made.
            frmCanvas.PSet (sngXIncr, sngYIncr)
            sngXIncr = sngXIncr + sngSlope
        Next sngYIncr
        
    End If

    frmCanvas.PSet (sngOriginalX, sngOriginalY)

End Sub
Private Function SwapEndPoints(Xstart As Single, Ystart As Single, Xend As Single, Yend As Single)

    ' This sub-routine is a helper for 'mdrLine' elsewhere within this module.
    Dim tempValue As Single
    
    ' Swap X
    tempValue = Xstart
    Xstart = Xend
    Xend = tempValue
    
    ' Swap Y
    tempValue = Ystart
    Ystart = Yend
    Yend = tempValue
    
End Function

Public Sub Draw_Vertices2(CurrentObject() As mdr2DObject, CurrentPictureBox As PictureBox)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    CurrentPictureBox.DrawStyle = g_lngDrawStyle
    CurrentPictureBox.DrawMode = vbCopyPen
    CurrentPictureBox.DrawWidth = 1
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                CurrentPictureBox.ForeColor = RGB(.Red, .Green, .Blue)
                
                ' Loop through the Vertices
                For intVertexIndex = LBound(.Vertex) To UBound(.Vertex)
                    xPos = .WorldPos.x
                    yPos = -.WorldPos.y
                    CurrentPictureBox.PSet (xPos, yPos)
                Next intVertexIndex
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub

Public Sub Draw_Faces(CurrentObject() As mdr2DObject, CurrentForm As Form)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    CurrentForm.DrawStyle = g_lngDrawStyle
    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawWidth = 1
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                CurrentForm.ForeColor = RGB(.Red, .Green, .Blue)
                
                If .Caption = "DefaultPlayerAmmo" Then
                    xPos = .TVertex(0).x
                    yPos = .TVertex(0).y
                    CurrentForm.PSet (xPos, yPos)
                Else
                    For intFaceIndex = LBound(.Face) To UBound(.Face)
                        
                        For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                        
                            intVertexIndex = .Face(intFaceIndex)(intK)
                            xPos = .TVertex(intVertexIndex).x
                            yPos = .TVertex(intVertexIndex).y
                            
                            If LBound(.Face(intFaceIndex)) = UBound(.Face(intFaceIndex)) Then
    
                            Else
                            
                                ' Normal Face; move to first point, then draw to the others.
                                ' ==========================================================
                                If intK = LBound(.Face(intFaceIndex)) Then
                                    ' Move to first point
                                    CurrentForm.Line (xPos, yPos)-(xPos, yPos)
'                                    Call mdrLine(xPos, yPos, xPos, yPos)
                                Else
                                    ' Draw to point
                                    CurrentForm.Line -(xPos, yPos)
'                                    Call mdrLine(frmCanvas.CurrentX, frmCanvas.CurrentY, xPos, yPos)
                                End If
                                                            
                            End If
                            
                        Next intK
                    Next intFaceIndex
                End If ' Is DefaultPlayerAmmo?
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub

Public Sub Draw_VerticesOnly(CurrentObject() As mdr2DObject, CurrentForm As Form)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    CurrentForm.DrawStyle = g_lngDrawStyle
    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawWidth = 1
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                CurrentForm.ForeColor = RGB(.Red, .Green, .Blue)
                
                ' Loop through the Vertices
                For intVertexIndex = LBound(.Vertex) To UBound(.Vertex)
                    
                    xPos = .TVertex(intVertexIndex).x
                    yPos = .TVertex(intVertexIndex).y
                
                    CurrentForm.PSet (xPos, yPos)
                    
                Next intVertexIndex
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub


Private Sub DrawVector(Vector As mdrVector3, Red As Integer, Green As Integer, Blue As Integer)

    Dim tempVector As mdrVector3
    
    tempVector = Vector
    tempVector = Vec3MultiplyByScalar(tempVector, 4000)
    tempVector = MatrixMultiplyVector(g_matViewMapping, tempVector)
    
    Dim sngMidpointX As Single
    Dim sngMidpointY As Single
    
    With frmCanvas
        sngMidpointX = .ScaleWidth / 2
        sngMidpointY = .ScaleHeight / 2
        .DrawStyle = g_lngDrawStyle
        .DrawWidth = 1
        .ForeColor = RGB(Red, Green, Blue)
        frmCanvas.Line (sngMidpointX, sngMidpointY)-(tempVector.x, tempVector.y)
    End With
    
End Sub


