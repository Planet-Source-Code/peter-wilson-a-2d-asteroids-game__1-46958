Attribute VB_Name = "mInternalData"
Option Explicit


Public Function CreateRandomShapeAsteroid(Radius As Single) As mdr2DObject

    ' Draws a deformed circle, by adjusting the radius at random intervals around the circumference.
    
    Dim sngAngle As Single
    Dim sngAngleIncrement As Single
    Dim sngMaxRadiusVariation As Single
    Dim sngNewRadius As Single
    Dim intMinSegmentAngle As Integer
    Dim intMaxSegmentAngle As Integer
    Dim sngRadiusVariation As Single
    Dim sngWorldX As Single
    Dim sngWorldY As Single
    Dim sngRadians As Single
    Dim intVertexCount As Integer
    Dim intFaceCount As Integer
    Dim intN As Integer
    
    Dim sngMinSize As Single
    Dim sngMaxSize As Single
    Dim sngAvgSize As Single
    
    Dim varVertices As Variant
    
    With CreateRandomShapeAsteroid
    
        ' Reset sizes to large opposite values.
        sngMinSize = Radius * 100
        sngMaxSize = -sngMinSize
        
        ' Set these Min/Max properties to make a random looking asteroid.
        ' Basically, this is a deformed circle.
        ' ===============================================================
        sngMaxRadiusVariation = Radius * 0.2 ' ie. 20% of Radius
        intMinSegmentAngle = 5
        intMaxSegmentAngle = 45
        
        ReDim .Vertex(0)
        intVertexCount = -1
        sngAngle = 0
        Do
        
            ' Get a new RND size (and remember the extreme sizes)
            sngNewRadius = GetRNDNumberBetween(Radius - sngMaxRadiusVariation, Radius + sngMaxRadiusVariation)
            If sngNewRadius < sngMinSize Then sngMinSize = sngNewRadius
            If sngNewRadius > sngMaxSize Then sngMaxSize = sngNewRadius
            sngAvgSize = sngAvgSize + sngNewRadius
            
            sngRadians = ConvertDeg2Rad(sngAngle)
            sngWorldX = Cos(sngRadians) * sngNewRadius
            sngWorldY = Sin(sngRadians) * sngNewRadius
            
            ' Create new Vertex
            intVertexCount = intVertexCount + 1
            ReDim Preserve .Vertex(intVertexCount)
            .Vertex(intVertexCount).x = sngWorldX
            .Vertex(intVertexCount).y = sngWorldY
            .Vertex(intVertexCount).w = 1
            
            sngAngleIncrement = GetRNDNumberBetween(intMinSegmentAngle, intMaxSegmentAngle)
            sngAngle = sngAngle + sngAngleIncrement
        
        Loop Until sngAngle >= 360
        
        .MinSize = sngMinSize
        .MaxSize = sngMaxSize
        .AvgSize = sngAvgSize / intVertexCount
        
        ReDim .TVertex(intVertexCount)
        
        ' Create the Asteroid's edges (ie. it's outer perimeter)
        ' ie. Face(0) = Array(0,1,2,...,n-1,n)
        ' =====================================================
        ReDim varVertices(intVertexCount + 1)
        ReDim .Face(0)
        For intN = 0 To intVertexCount
            varVertices(intN) = intN
        Next intN
        varVertices(intN) = 0
        .Face(0) = varVertices
        
        ' Create a Single Dot in the middle of the Asteroid and also create a face for it
        ' having only a single vertex.  This isn't really a face, more of a place-holder so
        ' I don't have to re-write my drawing routine.
        ' =================================================================================
        intVertexCount = UBound(.Vertex)
        ReDim Preserve .Vertex(intVertexCount + 1)
        ReDim Preserve .TVertex(intVertexCount + 1)
        .Vertex(intVertexCount + 1).x = 0
        .Vertex(intVertexCount + 1).y = 0
        .Vertex(intVertexCount + 1).w = 1
        
        intFaceCount = UBound(.Face)
        ReDim Preserve .Face(intFaceCount + 1)
        .Face(intFaceCount + 1) = Array(intVertexCount + 1)
    
    End With

End Function


Public Function Create_TestTriangle() As mdr2DObject

    With Create_TestTriangle
    
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        .Vertex(0).x = 0: .Vertex(0).y = 0: .Vertex(0).w = 1
        .Vertex(1).x = 0: .Vertex(1).y = 1: .Vertex(1).w = 1
        .Vertex(2).x = 2: .Vertex(2).y = 2: .Vertex(2).w = 1
        
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 0)
        
    End With
    
End Function

Public Function Create_ThrustFlame(LengthOfFlame As Single) As mdr2DObject

    With Create_ThrustFlame
    
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        .Vertex(0).x = -2: .Vertex(0).y = -1: .Vertex(0).w = 1
        .Vertex(1).x = 0: .Vertex(1).y = 0: .Vertex(1).w = 1
        .Vertex(2).x = 2: .Vertex(2).y = -1: .Vertex(2).w = 1
        .Vertex(3).x = 0: .Vertex(3).y = -1 - LengthOfFlame: .Vertex(3).w = 1
        
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 0)
        
    End With
    
End Function


Public Function Create_StarCluster(NumberOfStars As Integer, GameWorld As mdrWindow) As mdr2DObject

    ' Create a cluster of stars, and call this cluster a single 2D-Object
    
    With Create_StarCluster
    
        Dim intN As Integer
        ReDim .Vertex(NumberOfStars - 1)
        ReDim .TVertex(NumberOfStars - 1)
        
        For intN = 0 To (NumberOfStars - 1)
            .Vertex(intN).x = GetRNDNumberBetween(GameWorld.xMin, GameWorld.xMax)
            .Vertex(intN).y = GetRNDNumberBetween(GameWorld.yMin, GameWorld.yMax)
            .Vertex(intN).w = 1
        Next intN
                
        ' No Faces.
        ' =========
        ' Stars are just little dots.. ie. Vertices.  There is no need for a
        ' star cluster to have faces, so don't even define any.
        
    End With
    
End Function

Public Function Create_PlayerSpaceShip() As mdr2DObject

    With Create_PlayerSpaceShip
    
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        .Vertex(0).x = 0: .Vertex(0).y = 3: .Vertex(0).w = 1
        .Vertex(1).x = 3: .Vertex(1).y = -2: .Vertex(1).w = 1
        .Vertex(2).x = 0: .Vertex(2).y = 0: .Vertex(2).w = 1
        .Vertex(3).x = -3: .Vertex(3).y = -2: .Vertex(3).w = 1
        
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 0)
        
    End With
    
End Function

Public Function Create_PlayerAmmo() As mdr2DObject

    With Create_PlayerAmmo
    
        ReDim .Vertex(0)
        ReDim .TVertex(0)
        
        ' Make Ammo's equal to the gun-turret on the Player's space ship (optional - but hey... why not?!)
        .Vertex(0).x = 0: .Vertex(0).y = 3: .Vertex(0).w = 1
        
        ' No Faces.
        ' =========
        ' The default player ammo is just little dots.. ie. Vertices.  There is no need for
        ' the default player ammo to have faces, so don't even define any.
        ' Note: Some of the other, more exotic ammo will contain faces.... like heat-seeking-rockets! Whoosh!
        
    End With
    
End Function
Public Function Create_EnemySpaceShip1() As mdr2DObject

    With Create_EnemySpaceShip1
    
        ReDim .Vertex(7)
        ReDim .TVertex(7)
        
        .Vertex(0).w = 1
        .Vertex(1).w = 1
        .Vertex(2).w = 1
        .Vertex(3).w = 1
        .Vertex(4).w = 1
        .Vertex(5).w = 1
        .Vertex(6).w = 1
        .Vertex(7).w = 1
        
        .Vertex(0).x = 1: .Vertex(0).y = 2
        .Vertex(1).x = 2: .Vertex(1).y = 1
        .Vertex(2).x = 4: .Vertex(2).y = 0
        .Vertex(3).x = 2: .Vertex(3).y = -1
        .Vertex(4).x = -2: .Vertex(4).y = -1
        .Vertex(5).x = -4: .Vertex(5).y = 0
        .Vertex(6).x = -2: .Vertex(6).y = 1
        .Vertex(7).x = -1: .Vertex(7).y = 2
        
        
        ReDim .Face(1)
        .Face(0) = Array(1, 2, 3, 4, 5, 6, 1, 0, 7, 6)
        .Face(1) = Array(5, 2)
        
    End With
    
End Function
Public Function Create_Alphabet_A() As mdr2DObject

    With Create_Alphabet_A
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 2
        .Vertex(2).x = 2: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 2
        .Vertex(4).x = 4: .Vertex(4).y = 0
        
        ReDim .Face(1)
        .Face(0) = Array(0, 1, 2, 3)
        .Face(1) = Array(1, 3, 4)
        
    End With
    
End Function

Public Function Create_Alphabet_B() As mdr2DObject

    With Create_Alphabet_B
    
        Dim intN As Integer
        ReDim .Vertex(9)
        ReDim .TVertex(9)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 3
        .Vertex(2).x = 0: .Vertex(2).y = 6
        .Vertex(3).x = 3: .Vertex(3).y = 6
        .Vertex(4).x = 4: .Vertex(4).y = 5
        .Vertex(5).x = 4: .Vertex(5).y = 4
        .Vertex(6).x = 3: .Vertex(6).y = 3
        .Vertex(7).x = 4: .Vertex(7).y = 2
        .Vertex(8).x = 4: .Vertex(8).y = 1
        .Vertex(9).x = 3: .Vertex(9).y = 0
        
        ReDim .Face(1)
        .Face(0) = Array(0, 2, 3, 4, 5, 6, 1)
        .Face(1) = Array(6, 7, 8, 9, 0)
        
    End With
    
End Function

Public Function Create_Alphabet_C() As mdr2DObject

    With Create_Alphabet_C
    
        Dim intN As Integer
        ReDim .Vertex(7)
        ReDim .TVertex(7)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 1
        .Vertex(1).x = 3: .Vertex(1).y = 0
        .Vertex(2).x = 1: .Vertex(2).y = 0
        .Vertex(3).x = 0: .Vertex(3).y = 1
        .Vertex(4).x = 0: .Vertex(4).y = 5
        .Vertex(5).x = 1: .Vertex(5).y = 6
        .Vertex(6).x = 3: .Vertex(6).y = 6
        .Vertex(7).x = 4: .Vertex(7).y = 5
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4, 5, 6, 7)
        
    End With
    
End Function
Public Function Create_Alphabet_D() As mdr2DObject

    With Create_Alphabet_D
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 6
        .Vertex(2).x = 2: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 3
        .Vertex(4).x = 4: .Vertex(4).y = 1
        .Vertex(5).x = 3: .Vertex(5).y = 0
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4, 5, 0)
        
    End With
    
End Function

Public Function Create_Alphabet_E() As mdr2DObject

    With Create_Alphabet_E
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 0
        .Vertex(2).x = 0: .Vertex(2).y = 3
        .Vertex(3).x = 3: .Vertex(3).y = 3
        .Vertex(4).x = 0: .Vertex(4).y = 6
        .Vertex(5).x = 4: .Vertex(5).y = 6
        
        ReDim .Face(1)
        .Face(0) = Array(0, 1, 2)
        .Face(1) = Array(3, 2, 4, 5)
        
    End With
    
End Function

Public Function Create_Alphabet_F() As mdr2DObject

    With Create_Alphabet_F
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 3
        .Vertex(2).x = 3: .Vertex(2).y = 3
        .Vertex(3).x = 0: .Vertex(3).y = 6
        .Vertex(4).x = 4: .Vertex(4).y = 6
        
        ReDim .Face(1)
        .Face(0) = Array(0, 1, 2)
        .Face(1) = Array(1, 3, 4)
        
    End With
    
End Function

Public Function Create_Alphabet_G() As mdr2DObject

    With Create_Alphabet_G
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 2: .Vertex(0).y = 3
        .Vertex(1).x = 4: .Vertex(1).y = 3
        .Vertex(2).x = 4: .Vertex(2).y = 0
        .Vertex(3).x = 0: .Vertex(3).y = 0
        .Vertex(4).x = 0: .Vertex(4).y = 6
        .Vertex(5).x = 4: .Vertex(5).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4, 5)
        
    End With
    
End Function
Public Function Create_Alphabet_H() As mdr2DObject

    With Create_Alphabet_H
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 3
        .Vertex(2).x = 0: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 0
        .Vertex(4).x = 4: .Vertex(4).y = 3
        .Vertex(5).x = 4: .Vertex(5).y = 6
        
        ReDim .Face(2)
        .Face(0) = Array(0, 2)
        .Face(1) = Array(1, 4)
        .Face(2) = Array(3, 5)
        
    End With
    
End Function
Public Function Create_Alphabet_I() As mdr2DObject

    With Create_Alphabet_I
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 3: .Vertex(0).y = 0
        .Vertex(1).x = 2: .Vertex(1).y = 0
        .Vertex(2).x = 1: .Vertex(2).y = 0
        .Vertex(3).x = 3: .Vertex(3).y = 6
        .Vertex(4).x = 2: .Vertex(4).y = 6
        .Vertex(5).x = 1: .Vertex(5).y = 6
        
        ReDim .Face(2)
        .Face(0) = Array(0, 2)
        .Face(1) = Array(1, 4)
        .Face(2) = Array(3, 5)
        
    End With
    
End Function
Public Function Create_Alphabet_J() As mdr2DObject

    With Create_Alphabet_J
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 6
        .Vertex(1).x = 4: .Vertex(1).y = 2
        .Vertex(2).x = 3: .Vertex(2).y = 0
        .Vertex(3).x = 1: .Vertex(3).y = 0
        .Vertex(4).x = 0: .Vertex(4).y = 2
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4)
        
    End With
    
End Function
Public Function Create_Alphabet_K() As mdr2DObject

    With Create_Alphabet_K
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 3
        .Vertex(2).x = 0: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 0
        .Vertex(4).x = 4: .Vertex(4).y = 6
        
        ReDim .Face(1)
        .Face(0) = Array(0, 2)
        .Face(1) = Array(3, 1, 4)
        
    End With
    
End Function

Public Function Create_Alphabet_L() As mdr2DObject

    With Create_Alphabet_L
    
        Dim intN As Integer
        ReDim .Vertex(2)
        ReDim .TVertex(2)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 0
        .Vertex(2).x = 0: .Vertex(2).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2)
        
    End With
    
End Function

Public Function Create_Alphabet_M() As mdr2DObject

    With Create_Alphabet_M
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 6
        .Vertex(2).x = 2: .Vertex(2).y = 3
        .Vertex(3).x = 4: .Vertex(3).y = 6
        .Vertex(4).x = 4: .Vertex(4).y = 0
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4)
        
    End With
    
End Function

Public Function Create_Alphabet_N() As mdr2DObject

    With Create_Alphabet_N
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 6
        .Vertex(2).x = 4: .Vertex(2).y = 0
        .Vertex(3).x = 4: .Vertex(3).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3)
        
    End With
    
End Function

Public Function Create_Alphabet_O() As mdr2DObject

    With Create_Alphabet_O
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 6
        .Vertex(2).x = 4: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 0
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 0)
        
    End With
    
End Function

Public Function Create_Alphabet_P() As mdr2DObject

    With Create_Alphabet_P
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 3
        .Vertex(2).x = 0: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 6
        .Vertex(4).x = 4: .Vertex(4).y = 3
        
        ReDim .Face(0)
        .Face(0) = Array(0, 2, 3, 4, 1)
        
    End With
    
End Function

Public Function Create_Alphabet_Q() As mdr2DObject

    With Create_Alphabet_Q
    
        Dim intN As Integer
        ReDim .Vertex(6)
        ReDim .TVertex(6)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 6
        .Vertex(2).x = 4: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 2
        .Vertex(4).x = 2: .Vertex(4).y = 0
        .Vertex(5).x = 4: .Vertex(5).y = 0
        .Vertex(6).x = 2: .Vertex(6).y = 3
        
        ReDim .Face(1)
        .Face(0) = Array(0, 1, 2, 3, 4, 0)
        .Face(1) = Array(5, 6)
        
    End With
    
End Function
Public Function Create_Alphabet_R() As mdr2DObject

    With Create_Alphabet_R
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 3
        .Vertex(2).x = 0: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 6
        .Vertex(4).x = 4: .Vertex(4).y = 3
        .Vertex(5).x = 4: .Vertex(5).y = 0
        
        ReDim .Face(0)
        .Face(0) = Array(0, 2, 3, 4, 1, 5)
        
    End With
    
End Function
Public Function Create_Alphabet_S() As mdr2DObject

    With Create_Alphabet_S
    
        Dim intN As Integer
        ReDim .Vertex(5)
        ReDim .TVertex(5)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 0: .Vertex(0).y = 0
        .Vertex(1).x = 4: .Vertex(1).y = 0
        .Vertex(2).x = 4: .Vertex(2).y = 3
        .Vertex(3).x = 0: .Vertex(3).y = 3
        .Vertex(4).x = 0: .Vertex(4).y = 6
        .Vertex(5).x = 4: .Vertex(5).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4, 5)
        
    End With
    
End Function

Public Function Create_Alphabet_T() As mdr2DObject

    With Create_Alphabet_T
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 2: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 6
        .Vertex(2).x = 2: .Vertex(2).y = 6
        .Vertex(3).x = 4: .Vertex(3).y = 6
        
        ReDim .Face(1)
        .Face(0) = Array(0, 2)
        .Face(1) = Array(1, 3)
        
    End With
    
End Function
Public Function Create_Alphabet_U() As mdr2DObject

    With Create_Alphabet_U
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 6
        .Vertex(1).x = 4: .Vertex(1).y = 0
        .Vertex(2).x = 0: .Vertex(2).y = 0
        .Vertex(3).x = 0: .Vertex(3).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3)
        
    End With
    
End Function

Public Function Create_Alphabet_V() As mdr2DObject

    With Create_Alphabet_V
    
        Dim intN As Integer
        ReDim .Vertex(2)
        ReDim .TVertex(2)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 6
        .Vertex(1).x = 2: .Vertex(1).y = 0
        .Vertex(2).x = 0: .Vertex(2).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2)
        
    End With
    
End Function

Public Function Create_Alphabet_W() As mdr2DObject

    With Create_Alphabet_W
    
        Dim intN As Integer
        ReDim .Vertex(4)
        ReDim .TVertex(4)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 6
        .Vertex(1).x = 4: .Vertex(1).y = 0
        .Vertex(2).x = 2: .Vertex(2).y = 3
        .Vertex(3).x = 0: .Vertex(3).y = 0
        .Vertex(4).x = 0: .Vertex(4).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3, 4)
        
    End With
    
End Function

Public Function Create_Alphabet_X() As mdr2DObject

    With Create_Alphabet_X
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 6
        .Vertex(1).x = 0: .Vertex(1).y = 0
        .Vertex(2).x = 4: .Vertex(2).y = 0
        .Vertex(3).x = 0: .Vertex(3).y = 6
        
        ReDim .Face(1)
        .Face(0) = Array(0, 1)
        .Face(1) = Array(2, 3)
        
    End With
    
End Function

Public Function Create_Alphabet_Y() As mdr2DObject

    With Create_Alphabet_Y
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 6
        .Vertex(1).x = 2: .Vertex(1).y = 3
        .Vertex(2).x = 2: .Vertex(2).y = 0
        .Vertex(3).x = 0: .Vertex(3).y = 6
        
        ReDim .Face(1)
        .Face(0) = Array(0, 1, 2)
        .Face(1) = Array(1, 3)
        
    End With
    
End Function

Public Function Create_Alphabet_Z() As mdr2DObject

    With Create_Alphabet_Z
    
        Dim intN As Integer
        ReDim .Vertex(3)
        ReDim .TVertex(3)
        
        For intN = 0 To UBound(.Vertex)
            .Vertex(intN).w = 1
        Next intN
        
        .Vertex(0).x = 4: .Vertex(0).y = 0
        .Vertex(1).x = 0: .Vertex(1).y = 0
        .Vertex(2).x = 4: .Vertex(2).y = 6
        .Vertex(3).x = 0: .Vertex(3).y = 6
        
        ReDim .Face(0)
        .Face(0) = Array(0, 1, 2, 3)
        
    End With
    
End Function
