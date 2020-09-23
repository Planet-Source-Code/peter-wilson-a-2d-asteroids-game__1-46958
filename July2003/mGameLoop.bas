Attribute VB_Name = "mGameLoop"
Option Explicit

Public m_frmMap As frmMap

Public g_strGameState As String
Private m_lngApplicationState As Long

Public g_intCurrentLevel As Integer


' Stars in background
' ===================
Private m_StarsInCluster As Integer
Private m_StarsClusterMax As Integer
Private m_StarsInBackground() As mdr2DObject


Public m_PlayerShip() As mdr2DObject ' There's only a single player.
Private m_MaxAmmo As Integer
Public m_PlayerAmmo() As mdr2DObject ' PlayerAmmo is a special type of particle

Private m_Enemies() As mdr2DObject
Private m_GameObjects() As mdr2DObject

Private m_MaxAsteroids As Integer
Private m_Asteroids() As mdr2DObject


' Particles are objects that have a limited life span.
Private m_MaxParticles As Integer
Public m_Particles() As mdr2DObject

Private m_MaxVectorText As Integer
Private m_VectorText() As mdr2DObject

' Game World Window Limits ie. This is the game's world coordinates (which could be very large)
Private m_GameWorld As mdrWindow

' We will view the Game's world through the following window/s.
Public m_Window As mdrWindow

' Whatever we can see through the window, will be displayed on the viewport/s.
Private m_ViewPort As mdrWindow

Private m_Xmin As Single
Private m_Xmax As Single
Private m_Ymin As Single
Private m_Ymax As Single



' ViewPort Limits ie. Usually the limits of a VB form, or picturebox (which could be very small)
' Note: Just because your game's world is large, does not mean you need to display the whole world at
'       once. You can easily zoom in on some action just by changing the ViewPort values below.
'       Remember, the viewport is what you are looking at.
Private m_Umin As Single
Private m_Umax As Single
Private m_Vmin As Single
Private m_Vmax As Single


' Module Level Matrices (that don't change much)
Public m_matScale As mdrMATRIX3x3
Public g_lngDrawStyle As DrawStyleConstants
Public g_matViewMapping As mdrMATRIX3x3


Public g_blnDontClearScreen As Boolean
Private m_Alphabet(25) As mdr2DObject


' MIDI Volume (0=Off)
Public g_intMIDIVolume As Integer

Public g_blnUseMap As Boolean
Public g_blnExpertMode As Boolean
Public g_blnAllowRapidFire As Boolean
Public Function Create_Particles(Caption As String, NumberOfParticles As Integer, MinSize As Single, MaxSize As Single, WorldX As Single, WorldY As Single, VectorX As Single, VectorY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single, ZRotation As Single) As Integer

    ' "Attempts to create" the specified number of particles
    Create_Particles = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Particles(intN).Enabled = False Then
            ' This particle is no longer used, so we can use this one --> m_Particles(intN)
            
            If NumberOfParticles > 0 Then
                
                Select Case Caption
                    Case "Asteroid"
                        ' Create a random sized asteroid within the min/max parameters specified.
                        sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                        m_Particles(intN) = CreateRandomShapeAsteroid(sngRadius)
                
                    Case "Exhaust"
                        m_Particles(intN) = Create_ThrustFlame(6)
                    
                    Case "Exhaust_Blue"
                        m_Particles(intN) = Create_ThrustFlame(10)
                        
                    
                    Case "Exhaust_Smoke"
                        ' Create a random sized asteroid within the min/max parameters specified.
                        sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                        m_Particles(intN) = CreateRandomShapeAsteroid(sngRadius)
                        m_Particles(intN).ParticleMisc1 = 1
                        
                        Dim intRND As Integer
                        intRND = GetRNDNumberBetween(15, 70)
                        Call SendShortMIDIMessage(&H91, 30, 0) ' Turn off notes first
                        Call SendShortMIDIMessage(&H91, 20, 0)
                        Call SendShortMIDIMessage(&H91, 10, 0)

                        Call SendShortMIDIMessage(&H91, 30, g_intMIDIVolume * 0.5)
                        Call SendShortMIDIMessage(&H91, 20, g_intMIDIVolume * 0.6)
                        Call SendShortMIDIMessage(&H91, 10, g_intMIDIVolume)
                                                
                    
                    Case "DefaultPlayerAmmo"
                        m_Particles(intN) = Create_PlayerAmmo
                        m_Particles(intN).SpinVector = 0
                        m_Particles(intN).Vector.x = VectorX
                        m_Particles(intN).Vector.y = VectorY
                        
                    Case Else
                        ' Create a random sized asteroid within the min/max parameters specified.
                        sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                        m_Particles(intN) = CreateRandomShapeAsteroid(sngRadius)
                        
                End Select
                
                ' Fill-in some properties.
                With m_Particles(intN)
                    .Enabled = True
                    .Caption = Caption
                    .ParticleLifeRemaining = LifeTime
                    .WorldPos.x = WorldX
                    .WorldPos.y = WorldY
                    .WorldPos.w = 1
                    
                    ' Initial Vector
                    If VectorX = 0 Then
                        .Vector.x = GetRNDNumberBetween(-2, 2)
                    Else
                        .Vector.x = VectorX
                    End If
                    If VectorY = 0 Then
                        .Vector.y = GetRNDNumberBetween(-2, 2)
                    Else
                        .Vector.y = VectorY
                    End If
                    .Vector.w = 1
                        
                    If ZRotation = 0 Then .SpinVector = GetRNDNumberBetween(-4, 4)
                
                    .RotationAboutZ = ZRotation
                    .Red = Red: .Green = Green: .Blue = Blue
                End With
                
                NumberOfParticles = NumberOfParticles - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxParticles) Or (NumberOfParticles = 0)
        
End Function
Public Function Create_Asteroids(ByVal Qty As Integer, MinSize As Integer, MaxSize As Integer, WorldX As Single, WorldY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single) As Integer

    ' "Attempts to create" the specified number of Asteroids,
    ' and returns the number of Asteroids "actually created".
    Create_Asteroids = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Asteroids(intN).Enabled = False Then
            ' This Asteroid is no longer used, so we can use this one --> m_Asteroids(intN)
            
            If Qty > 0 Then
                
                ' Create a random sized asteroid within the min/max parameters specified.
                sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                m_Asteroids(intN) = CreateRandomShapeAsteroid(sngRadius)
                
                ' Reset the properties of this Asteroid.
                With m_Asteroids(intN)
                    .Enabled = True
                    .Caption = "Asteroid"
                    .Health = 100
                    .ParticleLifeRemaining = LifeTime
                    
                    ' Set a random starting position
                    Dim sngRadians As Single
                    sngRadians = ConvertDeg2Rad(GetRNDNumberBetween(0, 359))
                    .WorldPos.x = Cos(sngRadians) * (m_GameWorld.xMax * 0.95)
                    .WorldPos.y = Sin(sngRadians) * (m_GameWorld.yMax * 0.95)
                    .WorldPos.w = 1
                    
                    ' Initial Vector (Direction and Magnitude) depending on the size of the Asteroid.
                    Dim sngTemp As Single
                    sngTemp = 2 * (20 / sngRadius)
                    
                    .Vector.x = GetRNDNumberBetween(-sngTemp, sngTemp)
                    .Vector.y = GetRNDNumberBetween(-sngTemp, sngTemp)
                    .Vector.w = 1
                    
                    .SpinVector = GetRNDNumberBetween(-2, 2)
                    .RotationAboutZ = 0
                    
                    .Red = Red: .Green = Green: .Blue = Blue
                    
                End With
                
                Qty = Qty - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxParticles) Or (Qty = 0)
        
End Function

Private Sub Create_StarClusters()

    Dim intBrightness As Integer
    ReDim m_StarsInBackground(3)
    
    m_StarsInBackground(0) = Create_StarCluster(m_StarsInCluster, m_GameWorld)
    With m_StarsInBackground(0)
        .Caption = "Star Cluster"
        .SpinVector = 0.01   ' Any value from 0(zero) and up!
        intBrightness = 92
        .Red = intBrightness
        .Green = intBrightness
        .Blue = intBrightness
        .Enabled = True
    End With
        
    m_StarsInBackground(1) = Create_StarCluster(m_StarsInCluster, m_GameWorld)
    With m_StarsInBackground(1)
        .Caption = "Star Cluster"
        .SpinVector = 0.013   ' Any value from 0(zero) and up!
        intBrightness = 127
        .Red = intBrightness
        .Green = intBrightness
        .Blue = intBrightness
        .Enabled = True
    End With
    
    m_StarsInBackground(2) = Create_StarCluster(m_StarsInCluster, m_GameWorld)
    With m_StarsInBackground(2)
        .Caption = "Star Cluster"
        .SpinVector = 0.016   ' Any value from 0(zero) and up!
        intBrightness = 191
        .Red = intBrightness
        .Green = intBrightness
        .Blue = intBrightness
        .Enabled = True
    End With
    
    m_StarsInBackground(3) = Create_StarCluster(m_StarsInCluster, m_GameWorld)
    With m_StarsInBackground(3)
        .Caption = "Star Cluster"
        .SpinVector = 0.02    ' Any value from 0(zero) and up!
        intBrightness = 255
        .Red = intBrightness
        .Green = intBrightness
        .Blue = intBrightness
        .Enabled = True
    End With
End Sub

Private Sub ExplodeEverything()

    Static s_lngCounter As Long
    Dim intH As Integer
    
    s_lngCounter = s_lngCounter + 1
    If s_lngCounter = 1 Then
        ' Destroy Everything - ha ha ha ha!
        For intH = LBound(m_Asteroids) To UBound(m_Asteroids)
            m_Asteroids(intH).Enabled = False
            Call Create_Particles("Asteroid", 10, CInt(m_Asteroids(intH).AvgSize / 2), CInt(m_Asteroids(intH).AvgSize) * 2, m_Asteroids(intH).WorldPos.x, m_Asteroids(intH).WorldPos.y, 0, 0, 255, 255, 127, 40, 0)
        Next intH
    End If
    
    If s_lngCounter = 200 Then
        ' Destroy Everything - ha ha ha ha!
        For intH = LBound(m_Enemies) To UBound(m_Enemies)
            If m_Enemies(intH).Enabled = True Then
                m_Enemies(intH).Enabled = False
                Call Create_Particles("Enemy", 20, CInt(m_Enemies(intH).AvgSize * 2), CInt(m_Enemies(intH).AvgSize) * 3, m_Enemies(intH).WorldPos.x, m_Enemies(intH).WorldPos.y, 0, 0, 255, 255, 127, 40, 0)
            End If
        Next intH
    End If
    

    If s_lngCounter > 300 Then
        s_lngCounter = 0
        g_strGameState = "LevelComplete"
    End If

End Sub

Private Sub DoPlayerDead()

    Static s_lngCounter As Long
    
    s_lngCounter = s_lngCounter + 1
    
    If s_lngCounter > 200 Then
        s_lngCounter = 0
        g_strGameState = "LevelComplete"
    End If
    
End Sub

Private Sub ExplodeEverythingAndQuit()

    Static s_lngCounter As Long
    Dim intH As Integer
    
    s_lngCounter = s_lngCounter + 1
    If s_lngCounter = 1 Then
        ' Destroy Everything - ha ha ha ha!
        For intH = LBound(m_Asteroids) To UBound(m_Asteroids)
            m_Asteroids(intH).Enabled = False
            Call Create_Particles("Asteroid", 5, CInt(m_Asteroids(intH).AvgSize / 2), CInt(m_Asteroids(intH).AvgSize), m_Asteroids(intH).WorldPos.x, m_Asteroids(intH).WorldPos.y, 0, 0, 255, 255, 127, 60, 0)
        Next intH
    End If
    
    If s_lngCounter = 100 Then
        ' Destroy Everything - ha ha ha ha!
'        For intH = LBound(m_Enemies) To UBound(m_Enemies)
'            If m_Enemies(intH).Enabled = True Then
'                m_Enemies(intH).Enabled = False
'                Call Create_Particles("Asteroid", 20, CInt(m_Enemies(intH).AvgSize * 2), CInt(m_Enemies(intH).AvgSize) * 3, m_Enemies(intH).WorldPos.x, m_Enemies(intH).WorldPos.y, 0, 0, 255, 255, 127, 40,0)
'            End If
'        Next intH
    End If
    
    ' Disable (thus hiding) the star systems. It could also look cool to fade them to black.... or some other effect; you decide.
    If s_lngCounter = 100 Then m_StarsInBackground(3).Enabled = False
    If s_lngCounter = 120 Then m_StarsInBackground(2).Enabled = False
    If s_lngCounter = 140 Then m_StarsInBackground(1).Enabled = False
    If s_lngCounter = 160 Then m_StarsInBackground(0).Enabled = False
    
    
    If s_lngCounter > 300 Then
        s_lngCounter = 0
        g_strGameState = "Quit"
    End If

End Sub
Private Sub GameIsPaused()

    Static s_lngCounter As Long
    Static s_blnFlipFlop As Boolean
    
    s_lngCounter = s_lngCounter + 1
    If s_lngCounter = (2 ^ 31) - 1 Then s_lngCounter = 0
    
    If (s_lngCounter Mod 100) = 0 Then
        
        s_blnFlipFlop = Not s_blnFlipFlop
    
        If s_blnFlipFlop = True Then
            frmCanvas.Caption = "* * * P A U S E D * * *"
        Else
            frmCanvas.Caption = "* * *   G A M E   * * *"
        End If
    
    End If
    
End Sub

Private Sub Init_Alphabet()

    m_Alphabet(0) = Create_Alphabet_A
    m_Alphabet(1) = Create_Alphabet_B
    m_Alphabet(2) = Create_Alphabet_C
    m_Alphabet(3) = Create_Alphabet_D
    m_Alphabet(4) = Create_Alphabet_E
    m_Alphabet(5) = Create_Alphabet_F
    m_Alphabet(6) = Create_Alphabet_G
    m_Alphabet(7) = Create_Alphabet_H
    m_Alphabet(8) = Create_Alphabet_I
    m_Alphabet(9) = Create_Alphabet_J
    m_Alphabet(10) = Create_Alphabet_K
    m_Alphabet(11) = Create_Alphabet_L
    m_Alphabet(12) = Create_Alphabet_M
    m_Alphabet(13) = Create_Alphabet_N
    m_Alphabet(14) = Create_Alphabet_O
    m_Alphabet(15) = Create_Alphabet_P
    m_Alphabet(16) = Create_Alphabet_Q
    m_Alphabet(17) = Create_Alphabet_R
    m_Alphabet(18) = Create_Alphabet_S
    m_Alphabet(19) = Create_Alphabet_T
    m_Alphabet(20) = Create_Alphabet_U
    m_Alphabet(21) = Create_Alphabet_V
    m_Alphabet(22) = Create_Alphabet_W
    m_Alphabet(23) = Create_Alphabet_X
    m_Alphabet(24) = Create_Alphabet_Y
    m_Alphabet(25) = Create_Alphabet_Z
    
End Sub

Private Sub Init_Game()

    ' ===================
    ' Master Game Options
    ' ===================
    ' Set to False to speed up the game.
    g_blnUseMap = False
    
    ' Moving the Player's Space Ship can occur in two modes, Expert and Classic.
    g_blnExpertMode = False
    
    ' Allow Player's Space Ship to rapid fire it's ammo.
    g_blnAllowRapidFire = True
        
    
    ' ================================================================
    ' Hide Mouse (by moving it to the far bottom right)
    ' This method causes less problems than actually hiding the mouse,
    ' although moving the mouse can confuse the user, so careful.
    ' ================================================================
    Call SetCursorPos(Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
    
        
    ' ============================================================
    ' Show the Form's needed for drawing the game (Canvas and Map)
    ' ============================================================
    frmCanvas.Show
    If g_blnUseMap = True Then
        If m_frmMap Is Nothing Then
            Set m_frmMap = New frmMap
            Load m_frmMap
            Call m_frmMap.Move(0, 0)
        End If
        m_frmMap.Show vbModeless, frmCanvas
    End If
    
    
    ' Initializes the random-number generator.
    Randomize
    
    
    g_intCurrentLevel = 0 ' If you want the User to start on Level 5, put 4 here.
    
        
    
    ' Init Stars
    ' ==========
    'm_StarsClusterMax = 3
    m_StarsInCluster = 100
    
    ' Player Ammo.
    m_MaxAmmo = 3
    ReDim m_PlayerAmmo(m_MaxAmmo - 1)
    
    ' Init Particles
    ' ==============
    m_MaxParticles = 64
    ReDim m_Particles(m_MaxParticles - 1)
    
'''    ' Init Vector Text
'''    ' ================
'''    m_MaxVectorText = 32
'''    ReDim m_VectorText(m_MaxVectorText - 1)
'''    Call Init_Alphabet
    
    
'    Call Init_MIDISounds
    g_strGameState = "LevelComplete"
    
End Sub

Private Sub Init_MIDISounds()

    On Error GoTo errTrap
    
    g_intMIDIVolume = 127
    
    Call OpenMIDI
    
    Call ChangeInstrument(0, MelodicTom)         '   Fire Button
    Call ChangeInstrument(1, Gunshot)            '   Explosions
    Call ChangeInstrument(2, Seashore)           '   Thrust Forwards
    Call ChangeInstrument(3, AcousticGrandPiano) '   Getting-too-close-to-Asteroid-sound
    Call ChangeInstrument(4, Lead2_Sawtooth)
    Call ChangeInstrument(5, Lead1_Square)

    ' Drums
    Call SendShortMIDIMessage(&H99, 84, 0)  ' Good Sound for showing a temporary score for shooting something
    
    Exit Sub
errTrap:
    Select Case Err.Number
        Case vbObjectError + 1001
            MsgBox "MIDI sound effects can not be used due to the following reason:" & vbCrLf & vbCrLf & _
                   "'" & Err.Description & "'", vbExclamation, "Init_MIDISounds"
                   
    End Select
    
    g_intMIDIVolume = 0
        
End Sub

Public Sub Main()

    ' ==========================================================================
    ' This routine get's called by a Timer Event regardless of what's happening.
    ' (Although you can have multiple Timer Controls, it tends to make programs
    '  disorganised and less predictable. By using only a single Timer control,
    '  I have very strict control over what occurs and when. This routine is
    '  actually a mini-"state machine"... well actually most computer programs
    '  are, but I digress... look them up, learn them, they are cool.)
    ' ==========================================================================
        
    Select Case g_strGameState
        Case ""
            Call Init_Game
            
        Case "PlayingLevel"
            Call PlayGame
            
        Case "PlayerDead"
            Call PlayGame
            Call DoPlayerDead
            
        Case "LevelComplete"
            ' User has finished a level. Increment and reset game data for the next level.
            ' ============================================================================
            g_intCurrentLevel = g_intCurrentLevel + 1
            Call LoadLevel(g_intCurrentLevel)
            g_strGameState = "PlayingLevel"
            
            
        Case "ExplodeEverything"
            Call PlayGame
            Call ExplodeEverything
            
            
        Case "Paused"
            Call GameIsPaused
            
            
        Case "ExplodeEverythingAndQuit"
            g_strGameState = "Quit"
'            Call PlayGame
'            Call ExplodeEverythingAndQuit
            
            
        Case "Quit"
            Call CloseMIDI
            frmCanvas.Timer_DoAnimation.Enabled = False
            Unload frmCanvas
            
    End Select
    
    Call ProcessKeyboardInput
    
End Sub

Private Sub LoadLevel(ByVal Level As Integer)

    Dim intN As Integer
    Dim sngRadius As Single
    
    ' Reset Global Scale
    m_matScale = MatrixScaling(1, 1)
    
    ' =============
    ' Star Clusters
    ' =============
    ReDim m_StarsInBackground(m_StarsClusterMax)
    Call Create_StarClusters
    Call Calculate_StarClusters(g_matViewMapping)
    
    
    ' ================
    ' Create Asteroids
    ' ================
    ' One Large Asteroid can be split into two medium asteroids,
    ' and then each of these medium ones, can be split again into smaller ones.
    m_MaxAsteroids = Level
    If m_MaxAsteroids <> 0 Then
        ReDim m_Asteroids(m_MaxAsteroids - 1)
        Call Create_Asteroids(Level, 1, 30, 0, 0, 0, 192, 192, 0)
    End If
    
    
    ' ==================
    ' Create Player Ship
    ' ==================
    ReDim m_PlayerShip(0)
    m_PlayerShip(0) = Create_PlayerSpaceShip
    With m_PlayerShip(0)
        .Caption = "Player1"
        .Enabled = True
        .WorldPos.x = 0
        .WorldPos.y = 0
        .WorldPos.w = 1
        .Red = 0
        .Green = 255
        .Blue = 255
        .Health = 100
    End With
    
''    ' =====================================================================
''    ' Create Enemies
''    ' This should be space ships, but I've just made them asteroids for now
''    ' =====================================================================
''    Level = 1
''    ReDim m_Enemies(Int(Level / 2))
''    For intN = 0 To Int(Level / 2)
''        m_Enemies(intN) = Create_EnemySpaceShip1
''
''        With m_Enemies(intN)
''            .Caption = "Enemy" & intN
''            .Enabled = True
''            .WorldPos.x = GetRNDNumberBetween(m_GameWorld.xMin, m_GameWorld.xMax)
''            .WorldPos.y = GetRNDNumberBetween(m_GameWorld.yMin, m_GameWorld.yMax)
''            .WorldPos.w = 1
''            .Vector.x = 0
''            .Vector.y = 0
''            .SpinVector = 0
''            .AvgSize = 200
''            .Red = 0: .Green = 255: .Blue = 255
''        End With
''    Next intN
    
    
    ' Vector Text
    ' ===========
'
'    Call PrintVectorText("THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG", 0, 0)
'
'
    
End Sub

Public Sub PlayGame()
    
    Call SetViewWindow
    
    Call Calculate_StarClusters(g_matViewMapping)
    Call Calculate_Asteroids(g_matViewMapping)
    Call Calculate_Player(g_matViewMapping)
    Call Calculate_Particles(g_matViewMapping)
    
    Call Refresh_GameScreen(frmCanvas)
    
End Sub

Public Sub Calculate_Asteroids(ViewMapping As mdrMATRIX3x3)

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    On Error GoTo errTrap
    
    
    For intN = LBound(m_Asteroids) To UBound(m_Asteroids)
        With m_Asteroids(intN)
            If .Enabled = True Then
            
                ' Apply the direction/magnitude vector to the world coordinates.
                ' ==============================================================
                .WorldPos.x = .WorldPos.x + .Vector.x
                .WorldPos.y = .WorldPos.y + .Vector.y
                
                ' Clamp world position values to the game's world coordinate system
                ' (ie. Wrap aseteroids around the game's world)
                If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
                If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
                If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
                If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)
                
                
                ' Setup a Translation matrix
                ' ==========================
                matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
                
                ' Apply the spin vector to the Asteroid's rotation value.
                ' =======================================================
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                
                ' Setup a Rotation matrix
                ' =======================
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                
                ' Multiply matrices in the correct order.
                ' =======================================
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, m_matScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, ViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
                ' Conditionally Compiled (see Project Properties)
                #If gcShowVectors = -1 Then
                    ' Transform the Direction/Speed Vector to screen space
                    ' Do this step, only if you wish to display this vector on the screen.
                    ' Displaying the vector on screen, is only useful for debugging/instructional purposes.
                    ' Remember, DO NOT rotate the Direction/Speed vector (try it, and see what happens)
                    matResult = MatrixIdentity
                    matResult = MatrixMultiply(matResult, m_matScale)
                    matResult = MatrixMultiply(matResult, matTranslate)
                    matResult = MatrixMultiply(matResult, ViewMapping)
                    .TVector = MatrixMultiplyVector(matResult, .Vector)
                #End If
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub


Public Sub Calculate_StarClusters(ViewMapping As mdrMATRIX3x3)

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    On Error GoTo errTrap
    
    ' ===============================================================================================
    ' ARE YOU DEBUGGING HERE? ARE YOU DEBUGGING HERE? ARE YOU DEBUGGING HERE? ARE YOU DEBUGGING HERE?
    ' ===============================================================================================
    ' An error will occur here when the program is first run, however it has been error-trapped
    ' so just set your error handling to "Break in Class Module" and continue.
    For intN = LBound(m_StarsInBackground) To UBound(m_StarsInBackground)
        With m_StarsInBackground(intN)
                
                ' Apply the spin vector to the Asteroid's rotation value.
                ' =======================================================
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                
                ' Setup a Rotation matrix
                ' =======================
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
                
                ' Multiply matrices in the correct order.
                ' =======================================
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, ViewMapping)
            
            For intJ = LBound(.Vertex) To UBound(.Vertex)
                .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
            Next intJ
            
            
        End With
    Next intN
    
    Exit Sub
errTrap:

End Sub

Public Sub Calculate_VectorText()

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    On Error GoTo errTrap
    
    
    For intN = LBound(m_VectorText) To UBound(m_VectorText)
        With m_VectorText(intN)
            If .Enabled = True Then
            
                ' Translate
                ' =========
                .WorldPos.x = .WorldPos.x + .Vector.x
                .WorldPos.y = .WorldPos.y + .Vector.y
                
'                If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
'                If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
'                If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
'                If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)
                
                matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
                
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                Dim matVectorTextScale As mdrMATRIX3x3
                
                matVectorTextScale = MatrixScaling(200, 250)
                
                ' Multiply matrices in the correct order.
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, matVectorTextScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, g_matViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
                ' Reduce Particle life
                .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                If .ParticleLifeRemaining < 1 Then .Enabled = False
                
                
                ' Fade to Dull Red, then to black.
                .Red = .Red - 3
                .Green = .Green - 3
                .Blue = .Blue - 3
                                
                ' Conditionally Compiled (see Project Properties)
                #If gcShowVectors = -1 Then
                    ' Transform the Direction/Speed Vector to screen space
                    ' Do this step, only if you wish to display this vector on the screen.
                    ' Displaying the vector on screen, is only useful for debugging/instructional purposes.
                    ' Remember, DO NOT rotate the Direction/Speed vector (try it, and see what happens)
                    matResult = MatrixIdentity
                    matResult = MatrixMultiply(matResult, m_matScale)
                    matResult = MatrixMultiply(matResult, matTranslate)
                    matResult = MatrixMultiply(matResult, g_matViewMapping)
                    .TVector = MatrixMultiplyVector(matResult, .Vector)
                #End If
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub

Public Sub Calculate_Particles(ViewMapping As mdrMATRIX3x3)

    ' Processes all Particles (Asteroids, Exhaust, Bullets, Smoke, Flames, Explosions, etc.)
    
    Dim intN As Integer
    Dim intJ As Integer
    Dim matCustomScale As mdrMATRIX3x3
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
        
    For intN = LBound(m_Particles) To UBound(m_Particles)
        
        matCustomScale = MatrixScaling(1, 1)

        With m_Particles(intN)
            If .Enabled = True Then
                    
                Select Case .Caption
                    Case "Asteroid"
                        ' Fade to Dull Red, then to black.
                        .Red = .Red - 2
                        .Green = .Green - 4
                        .Blue = .Blue - 4
                        
                        ' Reduce Particle life
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                        
                    Case "Exhaust"
                        .Red = .Red - 16
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                    
                    Case "Exhaust_Blue"
                        .Blue = .Blue - 16
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                        
                    Case "Exhaust_Smoke"
                        .ParticleMisc1 = .ParticleMisc1 + 0.4
                        ' Smoke has 16 steps (not that it really matters)
                        If .ParticleMisc1 > 10 Then
                            .Red = .Red - 16
                            .Green = .Green - 16
                            .Blue = .Blue - 16
                        ElseIf .ParticleMisc1 > 3 Then
                            .Red = .Red + 24
                            .Green = .Green + 24
                            .Blue = .Blue + 24
                        End If
                        matCustomScale = MatrixScaling(.ParticleMisc1, .ParticleMisc1)
                        
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                    
                    Case "DefaultPlayerAmmo"
'                        .Red = .Red - 2
'                        .Green = .Green - 4
'                        .Blue = .Blue - 4
                        
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                    
                End Select
                
                ' Translate
                ' =========
                .WorldPos.x = .WorldPos.x + .Vector.x
                .WorldPos.y = .WorldPos.y + .Vector.y
                
                If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
                If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
                If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
                If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)
                
                matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
                
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                ' Multiply matrices in the correct order.
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, matCustomScale)
                matResult = MatrixMultiply(matResult, m_matScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, ViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
                Select Case .Caption
                    Case "DefaultPlayerAmmo"
'                        .Red = .Red - 2
'                        .Green = .Green - 6
'                        .Blue = .Blue - 6
                        
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
'                        If .ParticleLifeRemaining < 1 Then
'                            Call Create_Particles("Exhaust_Smoke", 1, 4, 4, .WorldPos.X, .WorldPos.Y, .Vector.X, .Vector.Y, 255, 0, 0, 12, 0)
'                        End If
                End Select
            End If ' Is Enabled?
            
        End With
    Next intN

End Sub

Public Sub Calculate_Enemies_Part1(ViewMapping As mdrMATRIX3x3)

    Dim intN As Integer
    Dim intJ As Integer
    Dim matScale As mdrMATRIX3x3
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    ' Custom Scale Matrix for Enemy Space ships
    matScale = MatrixScaling(1, 1)
    
    For intN = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intN)
                        
            ' Translate
            ' =========
            .WorldPos.x = .WorldPos.x + .Vector.x
            .WorldPos.y = .WorldPos.y + .Vector.y
            
            If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
            If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
            If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
            If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)
            
            matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
            
            .RotationAboutZ = .RotationAboutZ + .SpinVector
            matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))

            
            ' Multiply matrices in the correct order.
            matResult = MatrixIdentity
            matResult = MatrixMultiply(matResult, matScale) ' Custom Scale
            matResult = MatrixMultiply(matResult, m_matScale) ' Global Scale
            matResult = MatrixMultiply(matResult, matRotationAboutZ)
            matResult = MatrixMultiply(matResult, matTranslate)
            matResult = MatrixMultiply(matResult, ViewMapping)
            
            For intJ = LBound(.Vertex) To UBound(.Vertex)
                .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
            Next intJ
            
        End With
    Next intN

End Sub

Public Sub Calculate_Player(ViewMapping As mdrMATRIX3x3)

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    With m_PlayerShip(0)
                  
         If .Enabled = True Then
         
            ' Is Player Dead?
            ' ===============
            If m_PlayerShip(0).Health < 0 Then
                m_PlayerShip(0).Enabled = False
                Call Create_Particles("Asteroid", 32, 2, 40, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, 0, 0, 255, 0, 0, 240, 0)
                g_strGameState = "PlayerDead"
            End If
            
            If m_PlayerShip(0).Health < 10 Then
                m_PlayerShip(0).Red = 255
                m_PlayerShip(0).Green = 0
                m_PlayerShip(0).Blue = 0
            ElseIf m_PlayerShip(0).Health < 33 Then
                m_PlayerShip(0).Red = 255
                m_PlayerShip(0).Green = 220
                m_PlayerShip(0).Blue = 0
            ElseIf m_PlayerShip(0).Health < 66 Then
                m_PlayerShip(0).Red = 255
                m_PlayerShip(0).Green = 220
                m_PlayerShip(0).Blue = 220
            Else
                m_PlayerShip(0).Red = 0
                m_PlayerShip(0).Green = 255
                m_PlayerShip(0).Blue = 255
            End If
         
         
             If Abs(.SpinVector) > 20 Then
                
                Call Create_Particles("Asteroid", 40, 0.1, 2, .WorldPos.x, .WorldPos.y, 0, 0, 255, 0, 0, 60, 0)
                .Enabled = False
                Exit Sub
             End If
             
            ' Translate (ie. Move) the player ship to the correct location
            ' ============================================================
            .WorldPos.x = .WorldPos.x + .Vector.x
            .WorldPos.y = .WorldPos.y + .Vector.y
            
            
            ' Wrap the object to the boundry of the Game's World
            ' ==================================================
            If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
            If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
            If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
            If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)

            
''            ' Bounce off the Game Boundry
''            ' ===========================
''            If .WorldPos.X > m_GameWorld.xMax Then .Vector.X = -.Vector.X
''            If .WorldPos.X < m_GameWorld.xMin Then .Vector.X = -.Vector.X
''            If .WorldPos.Y > m_GameWorld.yMax Then .Vector.Y = -.Vector.Y
''            If .WorldPos.Y < m_GameWorld.yMin Then .Vector.Y = -.Vector.Y
            
            
            ' Build the Translation Matrix
            ' =============================
            matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
            
            ' Apply the spin vector to the Asteroid's rotation value.
            ' =======================================================
            .RotationAboutZ = .RotationAboutZ + .SpinVector
            
            ' Build the Rotation Matrix
            ' =========================
            matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
            
            ' Multiply the Matrices together (in the correct order)
            ' =====================================================
                    Dim matTemp1 As mdrMATRIX3x3
                    Dim matTemp2 As mdrMATRIX3x3
                    
''                    ' Build Matrix 1
''                    matTemp1 = MatrixIdentity
''                    matTemp1 = MatrixMultiply(matTemp1, m_matScale)
''                    matTemp1 = MatrixMultiply(matTemp1, matRotationAboutZ)
''
''                    ' Build Matrix 2
''                    matTemp2 = MatrixIdentity
''                    matTemp2 = MatrixMultiply(matTemp2, matTranslate)
''                    'matTemp2 = MatrixMultiply(matTemp2, ViewMapping)
''
''                    ' matResult = Matrix1 * Matrix2
''                    '   or
''                    ' matResult = Scale * RotationAboutZ * Translate * ViewMapping
''                    matResult = MatrixMultiply(matTemp1, matTemp2)
                    
                    ' Build Matrix 1
                    matResult = MatrixIdentity
                    matResult = MatrixMultiply(matResult, m_matScale)
                    matResult = MatrixMultiply(matResult, matRotationAboutZ)
                    matResult = MatrixMultiply(matResult, matTranslate)
                    matResult = MatrixMultiply(matResult, ViewMapping)
                    
            ' Multiply the Vertex points with the result Matrix.
            ' ==================================================
            For intJ = LBound(.Vertex) To UBound(.Vertex)
                .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
            Next intJ
            
            
        End If
    End With

End Sub
Public Sub Calculate_Enemies_Part2()

    ' Display Normalized Compass of Asteroids and recommends avoidance behaviour.

    Dim intEnemy As Integer
    Dim intAsteroid As Integer
    Dim tempV As mdrVector3
    Dim tempV3 As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    Dim sngMultiplier As Single
    
    On Error GoTo errTrap
    
    For intEnemy = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intEnemy)
            If .Enabled = True Then
                For intAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                    If m_Asteroids(intAsteroid).Enabled = True Then
                    
                    tempV.x = .WorldPos.x - m_Asteroids(intAsteroid).WorldPos.x
                    tempV.y = .WorldPos.y - m_Asteroids(intAsteroid).WorldPos.y
                    tempV.w = 1
                    
                    sngDistance = Vec3Length(tempV)
                    tempV = Vec3Normalize(tempV)
                    
                    #If gcShowVectors = -1 Then
                        VDisplay = tempV
                        VDisplay = Vec3MultiplyByScalar(VDisplay, 1900)
                        VDisplay.x = .WorldPos.x + VDisplay.x
                        VDisplay.y = .WorldPos.y + VDisplay.y
                    #End If

                    sngMultiplier = (1 - (sngDistance / m_GameWorld.xMax)) * 800
                    tempV = Vec3MultiplyByScalar(tempV, sngMultiplier)
                    
                    .Vector.x = tempV.x
                    .Vector.y = tempV.y
                    .Vector.w = 1
                    
                    #If gcShowVectors = -1 Then
                        VDisplay = MatrixMultiplyVector(g_matViewMapping, VDisplay)
                        frmCanvas.ForeColor = RGB(255, 255, 0)
                        frmCanvas.DrawWidth = 2
                        frmCanvas.PSet (VDisplay.x, VDisplay.y)
                        frmCanvas.Print tempV.x
                        
                    #End If

                    
                    ' =====================================================================
                    ' This is a VERY fun place to change paramters!
                    ' Minuses / Pluses. world coordininates, local, it just doesn't matter!
                    ' =====================================================================
                    If (sngDistance < 3000) Then
                        Call Create_Particles("Asteroid", 5, CInt(.AvgSize) * 2, CInt(.AvgSize) * 3, .WorldPos.x, .WorldPos.y, 0, 0, 255, 255, 127, 40, 0)
                        Call Create_Particles("Asteroid", 5, CInt(.AvgSize) * 2, CInt(.AvgSize) * 3, .WorldPos.x, .WorldPos.y, 0, 0, .Red, .Green, .Blue, 40, 0)
                        
                        .Enabled = False
                        
                        Dim intH As Integer
                        If (.Green > 127) And (.Blue > 127) Then
                            g_strGameState = "ExplodeEverything"
                        End If
                        
                    ElseIf (sngDistance < 6000) Then
                        tempV3 = Vec3Normalize(tempV)
                        tempV3 = Vec3MultiplyByScalar(tempV3, 1900)
'                        Call Create_Particles(1, 100, 100, .WorldPos.x, .WorldPos.y, tempV3.x, tempV3.y, 255, 255, 0, 10,0)
'                        Call Create_Particles(1, 100, 100, .WorldPos.x, .WorldPos.y, 0, 0, 255, 255, 0, 10,0)
                        Call Create_Particles("Asteroid", 1, CInt(.AvgSize), CInt(.AvgSize), .WorldPos.x, .WorldPos.y, 0, 0, .Red, .Green, .Blue, 15, 0)
                        .Red = .Red - 6
                        .Green = .Green + 3
                        .Blue = .Blue + 8
                    End If
                    
                    
                    End If ' Is Asteroid Enabled?
                Next intAsteroid
            End If ' Is Enemy Enabled?
        End With
    Next intEnemy
    
    Exit Sub
errTrap:
    
End Sub
Public Sub Calculate_EnemyAI(ViewMapping As mdrMATRIX3x3, CurrentForm As Form)

    Dim intEnemy As Integer
    Dim intAsteroid As Integer
    
    Dim V1 As mdrVector3
    Dim V2 As mdrVector3
    
    Dim WorldVectU As mdrVector3
    Dim WorldVectV As mdrVector3
    Dim sngWorldDotProduct As Single
    
    Dim sngETASeconds As Single
    
    ' Debug Display Graphics Only
    #If gcShowVectors = -1 Then
        Dim vectAsteroid As mdrVector3
        Dim vectDisplay As mdrVector3
        Dim strMsg As String
    #End If
    
    
    On Error GoTo errTrap
    
    For intEnemy = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intEnemy)
            If .Enabled = True Then
                
                For intAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                    If m_Asteroids(intAsteroid).Enabled = True Then

                        
                        ' v1 is difference between Enemy and Asteroid
                        V1 = Vect3Subtract(m_Asteroids(intAsteroid).WorldPos, .WorldPos)
                        V2 = Vect3Subtract(m_Asteroids(intAsteroid).Vector, .Vector)
                        
                        ' Normalize the vectors.
                        WorldVectU = Vec3Normalize(V2)
                        WorldVectV = Vec3Normalize(V1)
                        
                        ' Get the DotProduct between the two vectors.
                        sngWorldDotProduct = DotProduct(WorldVectU, WorldVectV)
                        
                        ' Threat Levels
                        .PreviousThreatLevel = .CurrentThreatLevel
                        .CurrentThreatLevel = sngWorldDotProduct
                        
                        #If gcShowVectors = -1 Then
                            strMsg = ""
                            
                            If sngWorldDotProduct < 6 Then
                                vectDisplay = Vec3Normalize(V1)
                                vectDisplay = Vec3MultiplyByScalar(vectDisplay, 7)
                                vectDisplay = Vect3Addition(vectDisplay, .WorldPos)
                                vectDisplay = MatrixMultiplyVector(ViewMapping, vectDisplay)
                                
                                vectAsteroid = m_Asteroids(intAsteroid).WorldPos
                                vectAsteroid = MatrixMultiplyVector(ViewMapping, vectAsteroid)
                                
                                CurrentForm.Font = "Arial Narrow"
                                CurrentForm.FontSize = 14
                                CurrentForm.DrawWidth = 1
                                CurrentForm.DrawStyle = vbDot
                                
'                                Call Calculate_AvoidanceVector(intEnemy, intAsteroid, ViewMapping, CurrentForm)
                                
                                strMsg = ""
                                If sngWorldDotProduct > 0.7 Then
                                    strMsg = "very safe - moving away"
                                    CurrentForm.ForeColor = RGB(0, 255, 0)
                                ElseIf sngWorldDotProduct > 0 Then
                                    strMsg = "safe - moving away"
                                    CurrentForm.ForeColor = RGB(255, 255, 0)
                                ElseIf sngWorldDotProduct < -0.98 Then
                                    strMsg = "Collision Danger!!!"
                                    CurrentForm.DrawStyle = vbSolid
                                    CurrentForm.DrawWidth = 3
                                    CurrentForm.ForeColor = RGB(255, 0, 0)
                                ElseIf sngWorldDotProduct < -0.98 Then
                                    strMsg = "danger"
                                    CurrentForm.DrawStyle = vbSolid
                                    CurrentForm.DrawWidth = 2
                                    CurrentForm.ForeColor = RGB(220, 64, 0)
                                ElseIf sngWorldDotProduct < -0.95 Then
                                    strMsg = "threat"
                                    CurrentForm.DrawStyle = vbSolid
                                    CurrentForm.ForeColor = RGB(255, 64, 0)
                                ElseIf sngWorldDotProduct < -0.8 Then
                                    strMsg = "possible threat"
                                    CurrentForm.DrawStyle = vbSolid
                                    CurrentForm.ForeColor = RGB(255, 127, 0)
                                ElseIf sngWorldDotProduct < 0 Then
                                    strMsg = "caution"
                                    CurrentForm.ForeColor = RGB(255, 255, 0)
                                End If
                                
                                strMsg = "DP" & Format(.CurrentThreatLevel, "0.0")
                                
'                                If (.CurrentThreatLevel < .PreviousThreatLevel) Then
'                                    strMsg = strMsg & "+"
'                                Else
'                                    strMsg = strMsg & "-"
'                                End If

                                CurrentForm.Line (vectAsteroid.x, vectAsteroid.y)-(vectDisplay.x, vectDisplay.y)

'                                strMsg = "Dot Product: " & Format(sngWorldDotProduct, "0.00") & " " & _
                                         "Dist.: " & Format(Vec3Length(V1), "0.00") & " " & _
                                         "Closing: " & Format(Vec3Length(V2), "0.00")
'                                CurrentForm.Print strMsg

                            End If
                                                        
                        #End If
                    
                    
                    End If ' Is Asteroid Enabled?
                Next intAsteroid
            End If ' Is Enemy Enabled?
        End With
    Next intEnemy
    
    Exit Sub
errTrap:
    
End Sub


Public Sub Calculate_Enemies_Part2b()

    ' Display Normalized Compass of other Ememies and recommends avoidance behaviour.

    Dim intEnemy As Integer
    Dim intOtherEnemy As Integer
    Dim tempV As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    
    Dim sngMultiplier As Single
        
    For intEnemy = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intEnemy)
            If .Enabled = True Then
                For intOtherEnemy = LBound(m_Enemies) To UBound(m_Enemies)
                    If (m_Enemies(intOtherEnemy).Enabled = True) And (intOtherEnemy <> intEnemy) Then
                    
                        tempV.x = .WorldPos.x - m_Enemies(intOtherEnemy).WorldPos.x
                        tempV.y = .WorldPos.y - m_Enemies(intOtherEnemy).WorldPos.y
                        tempV.w = 1
                        
                        sngDistance = Vec3Length(tempV)
                        If (sngDistance > 1) Then
                            tempV = Vec3Normalize(tempV)
                            
                            #If gcShowVectors = -1 Then
                                VDisplay = tempV
                                VDisplay = Vec3MultiplyByScalar(VDisplay, 1500)
                                VDisplay.x = .WorldPos.x + VDisplay.x
                                VDisplay.y = .WorldPos.y + VDisplay.y
                            #End If
                            
                            sngMultiplier = (1 - (sngDistance / m_GameWorld.xMax)) * 1500 ' <<< Change this * 400 bit!!!
                            tempV = Vec3MultiplyByScalar(tempV, sngMultiplier)
                            
'                            .Vector.x = tempV.x
'                            .Vector.y = tempV.y
'                            .Vector.w = 1
                            
                            #If gcShowVectors = -1 Then
                                VDisplay = MatrixMultiplyVector(g_matViewMapping, VDisplay)
                                frmCanvas.ForeColor = RGB(0, 255, 0)
                                frmCanvas.DrawWidth = 2
                                frmCanvas.PSet (VDisplay.x, VDisplay.y)
                            #End If
                        End If ' Is Asteroid close to us?
                                        
                    End If ' Is Other Enemy Enabled?
                Next intOtherEnemy
            End If ' Is Enemy Enabled?
        End With
    Next intEnemy
    

End Sub
Public Sub Calculate_Asteroids_Part2b()

    ' Display Normalized Compass of other Asteroids (optional)
    
    Dim intAsteroid As Integer
    Dim intOtherAsteroid As Integer
    Dim tempV As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    
    Dim sngMultiplier As Single
        
    For intAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
        With m_Asteroids(intAsteroid)
            If .Enabled = True Then
                For intOtherAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                    If (m_Asteroids(intOtherAsteroid).Enabled = True) And (intOtherAsteroid <> intAsteroid) Then
                    
                        tempV.x = .WorldPos.x - m_Asteroids(intOtherAsteroid).WorldPos.x
                        tempV.y = .WorldPos.y - m_Asteroids(intOtherAsteroid).WorldPos.y
                        tempV.w = 1
                        
                        sngDistance = Vec3Length(tempV)
                        tempV = Vec3Normalize(tempV)
                        
                        #If gcShowVectors = -1 Then
                            VDisplay = tempV
                            VDisplay = Vec3MultiplyByScalar(VDisplay, 3500)
                            VDisplay.x = .WorldPos.x + VDisplay.x
                            VDisplay.y = .WorldPos.y + VDisplay.y
                        #End If
    
                        sngMultiplier = (1 - (sngDistance / m_GameWorld.xMax)) * 10000 ' <<< Change this * 400 bit!!!
                        tempV = Vec3MultiplyByScalar(tempV, sngMultiplier)
                        
                        #If gcShowVectors = -1 Then
                            VDisplay = MatrixMultiplyVector(g_matViewMapping, VDisplay)
                            frmCanvas.ForeColor = RGB(0, 255, 255)
                            frmCanvas.DrawWidth = 2
                            frmCanvas.PSet (VDisplay.x, VDisplay.y)
                        #End If
                                        
                    End If ' Is Other Asteroid Enabled?
                Next intOtherAsteroid
            End If ' Is Asteroid Enabled?
        End With
    Next intAsteroid
    

End Sub
Public Sub Calculate_AvoidanceVector(EnemyIndex As Integer, AsteroidIndex As Integer, ViewMapping As mdrMATRIX3x3, CurrentForm As Form)

    ' Display Normalized Compass of other Asteroids (optional)
    
    Dim intAsteroid As Integer
    Dim intOtherAsteroid As Integer
    Dim tempV As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    
    Dim sngMultiplier As Single
    
    With m_Asteroids(AsteroidIndex)
        If .Enabled = True Then
        
            ' For each Asteroid, loop through all "other" Asteroids.
            For intOtherAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                If (m_Asteroids(intOtherAsteroid).Enabled = True) And (intOtherAsteroid <> intAsteroid) Then
                
                    tempV.x = .WorldPos.x - m_Asteroids(intOtherAsteroid).WorldPos.x
                    tempV.y = .WorldPos.y - m_Asteroids(intOtherAsteroid).WorldPos.y
                    tempV.w = 1
                    
                    tempV = Vec3MultiplyByScalar(tempV, -0.5)
                    tempV = Vect3Addition(tempV, .WorldPos)
                    
                    #If gcShowVectors = -1 Then
                    
                        tempV = MatrixMultiplyVector(g_matViewMapping, tempV)
                        CurrentForm.ForeColor = RGB(255, 0, 0)
                        CurrentForm.DrawWidth = 2
                        CurrentForm.PSet (tempV.x, tempV.y)
                    #End If
                                    
                End If ' Is Other Asteroid Enabled?
            Next intOtherAsteroid
        End If ' Is Asteroid Enabled?
    End With
    

End Sub


Public Sub Init_ViewMapping()

    ' The aspect ratio of most screen resolutions (ie. 1024x768 or 800x600) have an aspect ratio of 1.332 : 1
    ' Therefore I have made the Game's World coordinates slightly wider, so that everything looks square when
    ' you maximize the form.
    
    
    ' ===============================
    ' Set the size of the Game World.
    ' ===============================
    '   * The positive X axis points towards the right.
    '   * The positive Y axis points upwards to the top of the screen.
    '   * The positive Z axis points *into* the monitor. This is used for rotation.
    m_GameWorld.xMin = (-1000 * 1.333)
    m_GameWorld.xMax = (1000 * 1.333)
    m_GameWorld.yMin = -1000
    m_GameWorld.yMax = 1000
    
    
    ' Set the size of the window, through which we will view the Game world.
    ' (Change this window to scroll and zoom around the Game World)
    If (m_Window.xMin = m_Window.xMax) Then
        m_Window.xMin = m_GameWorld.xMin
        m_Window.xMax = m_GameWorld.xMax
        m_Window.yMin = m_GameWorld.yMin
        m_Window.yMax = m_GameWorld.yMax
    End If
    
    If (m_frmMap Is Nothing) = False Then
        With m_frmMap.pictMap
            .ScaleLeft = m_GameWorld.xMin
            .ScaleWidth = m_GameWorld.xMax - m_GameWorld.xMin
            .ScaleTop = m_GameWorld.yMin
            .ScaleHeight = (m_GameWorld.yMax - m_GameWorld.yMin)
        End With
        With m_frmMap
            .Shape1.Left = m_Window.xMin
            .Shape1.Top = -m_Window.yMax
            .Shape1.Width = m_Window.xMax - m_Window.xMin
            .Shape1.Height = m_Window.yMax - m_Window.yMin
        End With
    End If
    
    
    ' ==================================================================
    ' Set the size of the ViewPort (ie. normally a form, or picture box.
    ' ==================================================================
    '   This is normally set to the size of the form's internal drawing area (ie. ScaleWidth & ScaleHeight)
    m_ViewPort.xMin = 0
    m_ViewPort.xMax = frmCanvas.ScaleWidth
    m_ViewPort.yMin = frmCanvas.ScaleHeight
    m_ViewPort.yMax = 0
    
    
    ' ==========================
    ' Set the ViewMapping matrix
    ' ==========================
    g_matViewMapping = MatrixViewMapping(m_Window, m_ViewPort)
        
End Sub

Private Function PrintVectorText(Text As String, WorldX As Single, WorldY As Single) As Boolean

    Dim lngN As Long
    Dim strChar As String
    Dim intASCII As Integer
    Dim intAdjustedASCII As Integer
    Dim intTest As Integer
    
    Dim intXIncr As Integer
    Dim objCurrent As mdr2DObject
    Dim objNewText As mdr2DObject
    
    Dim intV As Integer
    Dim intA As Integer
    Dim intNewVertexCount As Integer
    Dim intNewFaceCount As Integer
    
    PrintVectorText = False
    If Text = "" Then Exit Function
    
    
    intNewVertexCount = 0
    intNewFaceCount = 0
    
    For lngN = 1 To Len(Text)
        strChar = Mid(Text, lngN, 1)
        intASCII = Asc(strChar)
        intAdjustedASCII = intASCII - 65
        If (intAdjustedASCII < 0 Or intAdjustedASCII > 25) = False Then
            
            objCurrent = m_Alphabet(intAdjustedASCII)
            With objCurrent
            
                ReDim Preserve objNewText.Vertex(intNewVertexCount + UBound(.Vertex))
                ReDim Preserve objNewText.TVertex(intNewVertexCount + UBound(.Vertex))
                
                ' Add new vertices from objCurrent to objNewText
                For intV = LBound(.Vertex) To UBound(.Vertex)
                    objNewText.Vertex(intV + intNewVertexCount) = .Vertex(intV)
                    objNewText.Vertex(intV + intNewVertexCount).x = objNewText.Vertex(intV + intNewVertexCount).x + intXIncr
                Next intV
                
                ReDim Preserve objNewText.Face(intNewFaceCount + UBound(.Face))
                ' Add new faces from objCurrent to objNewText
                For intV = LBound(.Face) To UBound(.Face)
                    objNewText.Face(intV + intNewFaceCount) = .Face(intV)
''                    For intA = LBound(objNewText.Face(intV + intNewFaceCount)) To UBound(objNewText.Face(intV + intNewFaceCount))
''                        objNewText.Face(intV + intNewFaceCount)(intA) = objNewText.Face(intV + intNewFaceCount)(intA) + intNewVertexCount
''                    Next intA
                Next intV
                intNewFaceCount = intNewFaceCount + UBound(.Face) + 2
                
                intNewVertexCount = intNewVertexCount + UBound(.Vertex) + 2
                intXIncr = lngN * 5
            End With
        End If
    Next lngN
    
    
    intTest = 0
    For intTest = 0 To UBound(m_VectorText)
        If m_VectorText(intTest).Enabled = False Then
            m_VectorText(intTest) = objNewText
             With m_VectorText(intTest)
                .Caption = "Letter: " & strChar
                .Enabled = True
                .ParticleLifeRemaining = 50
                .RotationAboutZ = 0
                .Vector.x = GetRNDNumberBetween(-20, 20): .Vector.y = GetRNDNumberBetween(-20, 20): .Vector.w = 1
                .SpinVector = GetRNDNumberBetween(-0.5, 0.5)
                .WorldPos.x = WorldX + (lngN * 1100)
                .WorldPos.y = WorldY
                .Red = 255
                .Green = 255
                .Blue = 255
             End With
             
            Exit For
        End If
    Next intTest
    
End Function

Private Sub Refresh_GameScreen(CurrentForm As Form)

    ' If I see one more game that uses BitBlt - I am going to Scream!  Arrrggghhhh!!!!!
    ' =================================================================================
    '   * You don't need BitBlt. I have absolutely no clue why people use it 99% of the time, when it is simply not needed.
    '   * You don't need DoEvents (Actually, this can cause more problems than it solves, so do yourself
    '     a big-big-BIG favour and just pretend it doesn't exist.)
    '   * You don't need to use Refresh (unless you want to slow down your program... which might be good for debugging)
    '     Set the form (or pictureboxes) AutoDraw property to True.
    '   * You don't need to use more than a single Timer control for your game.... really... you don't!
    '   * Try to learn a few API's, Particulary for drawing graphics and handling the mouse and keyboard input.
    '     They are not too hard once you get the hang of them.  Some are REALLY easy to use!
    
    
    ' Important: Clear the screen before drawing anything.
    ' ====================================================
    ' This may sound obvious to some, but if you don't then you'll
    ' need to use BitBlt (or something similar) which would be a bit stupid.
    ' You should always try to minimise the flicker in your game... only when this fails, should you use a Blittling Process.
    If g_blnDontClearScreen = False Then CurrentForm.Cls
    
    
    Call DrawCrossHairs(CurrentForm)
    Call Draw_VerticesOnly(m_StarsInBackground, CurrentForm)
    Call Draw_Faces(m_Asteroids, CurrentForm)
    Call Draw_Faces(m_Particles, CurrentForm)
    Call Draw_Faces(m_PlayerShip, CurrentForm)
    
    
    ' Draw stuff to the mini-map (using the quickest way possible)
    If (m_frmMap Is Nothing) = False Then
        If m_frmMap.Visible = True Then
            m_frmMap.pictMap.Cls
            Call Draw_Vertices2(m_Asteroids, m_frmMap.pictMap)
            Call Draw_Vertices2(m_PlayerShip, m_frmMap.pictMap)
            Call Draw_Vertices2(m_Particles, m_frmMap.pictMap)
        End If
    End If
    
End Sub

Private Sub SetViewWindow()
    
    Dim intAsteroidIndex As Integer
    Dim intIndex As Integer
    Dim sngClosestDistance As Single
    Dim sngDistance As Single
    
    
    ' Exit the Sub immediately, to cancel ALL zooming.
    ' ================================================
'    Exit Sub


    ' Set the ViewPort to a Fixed size, and don't perform any Zooming.
    ' ================================================================
'    m_Window.xMin = m_PlayerShip(0).WorldPos.X - (350 * 1.33)
'    m_Window.xMax = m_PlayerShip(0).WorldPos.X + (350 * 1.33)
'    m_Window.yMin = m_PlayerShip(0).WorldPos.Y - 350
'    m_Window.yMax = m_PlayerShip(0).WorldPos.Y + 350
'    Call Init_ViewMapping
'    Exit Sub
    
    
    
    ' Adjust the ViewPort (Zoom & Pan), depending on which Asteroid is the Closest
    ' (Doing all of this calculation slows down the Game quite a bit)
    ' ============================================================================
    intIndex = -1
    sngClosestDistance = (m_GameWorld.yMax - m_GameWorld.yMin) * 0.3 ' Set to far away.

    ' Loop through all Asteroids, looking for the closest one.
    For intAsteroidIndex = LBound(m_Asteroids) To UBound(m_Asteroids)
        If m_Asteroids(intAsteroidIndex).Enabled = True Then
            
            ' Find distance between current Asteroid and the Player's ship
            ' ============================================================
            sngDistance = Vec3Length(Vect3Subtract(m_Asteroids(intAsteroidIndex).WorldPos, m_PlayerShip(0).WorldPos))
            
            
            ' Do basic Collision detection (with the closest Asteroid)
            ' ========================================================
            If sngDistance < m_Asteroids(intAsteroidIndex).MinSize Then
                m_PlayerShip(0).Health = m_PlayerShip(0).Health - 33
            ElseIf sngDistance < m_Asteroids(intAsteroidIndex).AvgSize Then
                m_PlayerShip(0).Health = m_PlayerShip(0).Health - 2
                Call Create_Particles("Asteroid", 3, 0.5, 0.5, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, 0, 0, 255, 0, 0, 60, 0)
            ElseIf sngDistance < m_Asteroids(intAsteroidIndex).MaxSize Then
                m_PlayerShip(0).Health = m_PlayerShip(0).Health - 1
                Call Create_Particles("Asteroid", 3, 0.3, 0.3, m_PlayerShip(0).WorldPos.x, m_PlayerShip(0).WorldPos.y, 0, 0, 255, 196, 0, 20, 0)
            End If
            
            
            ' Reset color (optional)
            m_Asteroids(intAsteroidIndex).Red = 0
            m_Asteroids(intAsteroidIndex).Green = 128
            m_Asteroids(intAsteroidIndex).Blue = 128
        
            If Abs(m_Asteroids(intAsteroidIndex).Vector.x > 2) Or Abs(m_Asteroids(intAsteroidIndex).Vector.y > 2) Then
                ' Ignore Fast Moving Asteroids from Zoom calculations.
                m_Asteroids(intAsteroidIndex).Red = 0
                m_Asteroids(intAsteroidIndex).Green = 92
                m_Asteroids(intAsteroidIndex).Blue = 92
            Else
                If sngDistance < sngClosestDistance Then
                    sngClosestDistance = sngDistance
                    intIndex = intAsteroidIndex
                End If
            End If
            
            
        End If
        
    Next intAsteroidIndex
    
    If intIndex <> -1 Then
        m_Asteroids(intIndex).Red = 0
        m_Asteroids(intIndex).Green = 255
        m_Asteroids(intIndex).Blue = 255
    End If


    ' Produce heart beat sound, depending on how close we are to Asteroid.
    Static s_lngCounter As Long
    s_lngCounter = s_lngCounter + 4
    If s_lngCounter >= sngClosestDistance Then
        Call SendShortMIDIMessage(&H93, 30, 0)
        Call SendShortMIDIMessage(&H93, 30, g_intMIDIVolume * 0.3)
        s_lngCounter = 0
    End If
    
    ' Find the mid-point between the Closest Asteroid and the Player's space ship (then scale this value)
    Dim tempV As mdrVector3
    If intIndex <> -1 Then
        tempV = Vect3Subtract(m_Asteroids(intIndex).WorldPos, m_PlayerShip(0).WorldPos)
        tempV = Vec3MultiplyByScalar(tempV, 0.4) ' << Change this fractional part to between 0 and 1.
    End If


    ' Smoothly pan from one object to the next, by adjusting the camera panning values s_tempV.X & s_tempV.Y
    Static s_tempV As mdrVector3
    If s_tempV.x > tempV.x Then s_tempV.x = s_tempV.x - 1 ' <<< Adjust how smooth/quick the Pan should be.
    If s_tempV.x < tempV.x Then s_tempV.x = s_tempV.x + 1 ' <<< etc.
    If s_tempV.y > tempV.y Then s_tempV.y = s_tempV.y - 1
    If s_tempV.y < tempV.y Then s_tempV.y = s_tempV.y + 1
    
    
    ' Prevent extreme zoom values, by limiting to 150 (or any value higher than 0)
    If sngClosestDistance < 150 Then sngClosestDistance = 150
    
    
    ' This is the part that actually determins the visible area to be drawn.
    ' ======================================================================
    m_Window.xMin = m_PlayerShip(0).WorldPos.x - (sngClosestDistance * 1.33) + s_tempV.x
    m_Window.xMax = m_PlayerShip(0).WorldPos.x + (sngClosestDistance * 1.33) + s_tempV.x
    m_Window.yMin = m_PlayerShip(0).WorldPos.y - sngClosestDistance + s_tempV.y
    m_Window.yMax = m_PlayerShip(0).WorldPos.y + sngClosestDistance + s_tempV.y
    Call Init_ViewMapping
    
End Sub

