Attribute VB_Name = "mKeyboard"
Option Explicit

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub ProcessKeyboardInput()

    On Error GoTo errTrap
    
    Static lngCounter As Long
    lngCounter = lngCounter + 1
    If lngCounter > 2 ^ 31 - 1 Then lngCounter = 0
    
    Static s_strPreviousValue As String
    Static s_blnKeyDeBounce As Boolean
    Static s_blnKeyDeBounce_TAB As Boolean
    Static s_blnKeyDeBounce_FullScreen As Boolean
    
    ' Keyboard DeBounce
    Static s_lngKeyCombinations As Long
    
    Dim lngKeyCombinations As Long
    Dim lngKeyState As Long
    Dim sngSpeedIncrement As Single
    Dim sngSpeedMagnitude As Single
    Dim sngRadians As Single
    
    sngSpeedIncrement = 6
    
    lngKeyState = GetKeyState(vbKeyLeft)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 1 ' Left
    
    lngKeyState = GetKeyState(vbKeyRight)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 2 ' Right
        
    lngKeyState = GetKeyState(vbKeyUp)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 4 ' Thrust
    
    lngKeyState = GetKeyState(vbKeyDown)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 8 ' Cheat Mode - Dead Stop
    
    lngKeyState = GetKeyState(&HA2) ' Right Control Key
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 16 ' Fire!
    
    
    
    
    If g_blnExpertMode = True Then
        ' ===================================================================
        ' Rotate player ship using Expert mode (very hard, but more realistic)
        ' ===================================================================
        If (lngKeyCombinations And 1) = 1 Then m_PlayerShip(0).SpinVector = m_PlayerShip(0).SpinVector + (sngSpeedIncrement / 64)
        If (lngKeyCombinations And 2) = 2 Then m_PlayerShip(0).SpinVector = m_PlayerShip(0).SpinVector - (sngSpeedIncrement / 64)
    Else
        ' Rotate player ship using Classic Mode (easy and fun)
        ' ====================================================
        m_PlayerShip(0).SpinVector = 0 ' Reset the SpinVector
        If (lngKeyCombinations And 1) = 1 Then m_PlayerShip(0).SpinVector = sngSpeedIncrement
        If (lngKeyCombinations And 2) = 2 Then m_PlayerShip(0).SpinVector = -sngSpeedIncrement
    End If
    
    
    
    ' ===============
    ' Thrust Forwards
    ' ===============
    If (lngKeyCombinations And 4) = 4 Then
        
        ' ===================================================================
        ' Calculate a new thrust vector. Note: This is a pretty small vector.
        ' ===================================================================
        Dim ThrustVector As mdrVector3
        sngRadians = ConvertDeg2Rad(m_PlayerShip(0).RotationAboutZ + 90) ' This is offset by 90 degrees because I drew my ship's thrust port, -90 degrees the wrong way. Oh well! Too lazy to redefine the Player Space ship routine.
        ThrustVector.x = Cos(sngRadians) * 0.03 ' <<< Change this to increase engine thrust!
        ThrustVector.y = Sin(sngRadians) * 0.03 ' <<< Change this to increase engine thrust!
        ThrustVector.w = 1
        
        ' ================================================
        ' Add the Thrust Vector, to the PlayerShip vector.
        ' ================================================
        m_PlayerShip(0).Vector = Vect3Addition(ThrustVector, m_PlayerShip(0).Vector)
        
        
        ' =============================================================================================
        ' Draw some thrust and smoke, depending on the speed of the ship (ie. the length of the Vector)
        ' Don't forget, they don't get drawn here, just created here... Big difference!
        ' =============================================================================================
        sngSpeedMagnitude = Vec3Length(m_PlayerShip(0).Vector)
        If (lngCounter Mod 4) = 0 Then
            Call Create_Particles("Exhaust", 1, 1, 1, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, -ThrustVector.x * 64, -ThrustVector.y * 64, (255 + 16), 0, 0, 16, m_PlayerShip(0).RotationAboutZ)
        ElseIf (lngCounter Mod 4) = 1 Then
            Call Create_Particles("Exhaust_Blue", 1, 1, 1, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, -ThrustVector.x * 96, -ThrustVector.y * 96, 0, 0, (255 + 16), 16, m_PlayerShip(0).RotationAboutZ)
        ElseIf (lngCounter Mod 4) = 2 Then
            ' Only produce smoke, when the player's space ship is travelling fast.
            If sngSpeedMagnitude > 5.8 Then
                Call Create_Particles("Exhaust_Smoke", 1, 1, 1, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, -ThrustVector.x * 64, -ThrustVector.y * 64, 0, 0, 0, 40, 0)
            End If
        End If
        
        Call SendShortMIDIMessage(&H92, 9, g_intMIDIVolume * 0.3)
    Else
        Call SendShortMIDIMessage(&H92, 9, 0)
        
    End If
    
    
    ' ===================================================
    ' Dead Stop. This is a cheat and should be taken out.
    ' ===================================================
    If (lngKeyCombinations And 8) = 8 Then
        m_PlayerShip(0).Vector.x = 0
        m_PlayerShip(0).Vector.y = 0
        m_PlayerShip(0).SpinVector = 0
'        m_PlayerShip(0).RotationAboutZ = 0
    End If
    
    
    ' ================
    ' Fire Player Ammo
    ' ================
    If g_blnAllowRapidFire = True Then
        If (lngKeyCombinations And 16) = 16 Then
            Call SendShortMIDIMessage(&H99, 27, 0)
            Call SendShortMIDIMessage(&H99, 27, g_intMIDIVolume)
                Dim AmmoVector As mdrVector3
                sngRadians = ConvertDeg2Rad(m_PlayerShip(0).RotationAboutZ + 90)
                AmmoVector.x = Cos(sngRadians) * 3
                AmmoVector.y = Sin(sngRadians) * 3
                AmmoVector.w = 1
                ' Add a recoil to the PlayerShip. When they fire ammo, they go backwards a little
    ''            m_PlayerShip(0).Vector.X = m_PlayerShip(0).Vector.X - (AmmoVector.X / 20)
    ''            m_PlayerShip(0).Vector.Y = m_PlayerShip(0).Vector.Y - (AmmoVector.Y / 20)
                Call Create_Particles("DefaultPlayerAmmo", 1, 1, 1, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, (m_PlayerShip(0).Vector.x + AmmoVector.x), (m_PlayerShip(0).Vector.y + AmmoVector.y), 255, 255, 255, 180, m_PlayerShip(0).RotationAboutZ)
        End If
    Else
        If (lngKeyCombinations And 16 And Not s_lngKeyCombinations) = 16 Then
            Call SendShortMIDIMessage(&H99, 27, 0)
            Call SendShortMIDIMessage(&H99, 27, g_intMIDIVolume)
                sngRadians = ConvertDeg2Rad(m_PlayerShip(0).RotationAboutZ + 90)
                AmmoVector.x = Cos(sngRadians) * 3
                AmmoVector.y = Sin(sngRadians) * 3
                AmmoVector.w = 1
                ' Add a recoil to the PlayerShip. When they fire ammo, they go backwards a little
    ''            m_PlayerShip(0).Vector.X = m_PlayerShip(0).Vector.X - (AmmoVector.X / 20)
    ''            m_PlayerShip(0).Vector.Y = m_PlayerShip(0).Vector.Y - (AmmoVector.Y / 20)
                Call Create_Particles("DefaultPlayerAmmo", 1, 1, 1, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, (m_PlayerShip(0).Vector.x + AmmoVector.x), (m_PlayerShip(0).Vector.y + AmmoVector.y), 255, 255, 255, 180, m_PlayerShip(0).RotationAboutZ)
        End If
    End If ' Rapid Fire
    
    
    ' =======================
    ' DO NOT clear the screen
    ' =======================
    lngKeyState = GetKeyState(vbKeyC)
    If (lngKeyState And &H8000) Then g_blnDontClearScreen = True Else g_blnDontClearScreen = False
    
    
    ' Check the Space Bar for level complete. (Also shows how to "de-bounce" the space bar)
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        If s_blnKeyDeBounce = False Then
            s_blnKeyDeBounce = True
            g_strGameState = "LevelComplete"
        End If
    Else
        s_blnKeyDeBounce = False
    End If
    
    
    ' Check for Pause/Resume
    lngKeyState = GetKeyState(vbKeyP)
    If (lngKeyState And &H8000) Then
        If g_strGameState <> "Paused" Then
            s_strPreviousValue = g_strGameState
            g_strGameState = "Paused"
        Else
            g_strGameState = s_strPreviousValue
            frmCanvas.Caption = App.Comments
        End If
    End If
    
    
    ' Check for ESCAPE key.
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then g_strGameState = "ExplodeEverythingAndQuit"
    
    
    ' Check for TAB key to Hide/Show Map
    lngKeyState = GetKeyState(vbKeyTab)
    If (lngKeyState And &H8000) Then
        If s_blnKeyDeBounce_TAB = False Then
            s_blnKeyDeBounce_TAB = True
            m_frmMap.Visible = Not m_frmMap.Visible
        End If
    Else
        s_blnKeyDeBounce_TAB = False
    End If
    
    
        
    ' =========================================================
    ' Global Geometry Scale Value (optional)
    ' This is NOT a cosmetic change, this is a geometry change.
    ' =========================================================
    ' This is NOT really needed - you can take it out everywhere is appears,
    ' but what the heck, change it and see what happens anyways.
    lngKeyState = GetKeyState(vbKeyShift)
    If (lngKeyState And &H8000) Then
        m_matScale = MatrixScaling(8, 8) ' Magnify 8 Times!
        g_lngDrawStyle = vbDot
    Else
        g_lngDrawStyle = vbSolid
        m_matScale = MatrixScaling(1, 1) ' Normal size = 1,1
    End If
    
    
    s_lngKeyCombinations = lngKeyCombinations
    
    Exit Sub
errTrap:
    
End Sub

