Attribute VB_Name = "mDataStructures"
Option Explicit

' Used to define the Game's world, the Window through which we view the world
' and the viewport that will eventually display everything.
' GameWorld -> Window -> Viewport.
' ===========================================================================
Public Type mdrWindow
    xMin As Single
    xMax As Single
    yMin As Single
    yMax As Single
End Type


' This is a three dimensional vector, used to manipulate two dimensional objects.
' (Note: If this was a 3D graphics application, then we would use a 4D vector.)
' ===============================================================================
Public Type mdrVector3
    x As Single
    y As Single
    w As Single ' <<< w is not used anywhere in this application, and is always set to 1.0
                ' The only reason it is in here, is for future compatability.
                ' In other words, I might use w later to perform a tricky maths thing,
                ' but for now, it can be totally ignored (and if you want, you can even
                ' delete all references to it from this code)
End Type


' Stores Mathematical Transformation values.
' ==========================================
Public Type mdrMATRIX3x3
    rc11 As Single: rc12 As Single: rc13 As Single
    rc21 As Single: rc22 As Single: rc23 As Single
    rc31 As Single: rc32 As Single: rc33 As Single
    ' rc31 & rc32 will always equal 0, whilst rc33 always equals 1.
    ' (This is related to the 'w' parameter as mention above.)
End Type


' I use the following custom data type (mdr2DObject) for basically everything, including
' Asteroids, Player and Enemy Ships, Vector Text, Bullets, Smoke and Fire. As you can see
' this is a very versitile object.
' =======================================================================================
Public Type mdr2DObject
    
    ' General Properties
    Caption As String   ' (Optional)
    Enabled As Boolean  ' Normally TRUE, if FALSE then no calculations take place.
    
    ' Particle Counters
    ParticleLifeRemaining As Single ' A Particle Object is only Enabled for a short time.
    ParticleMisc1 As Single
    ParticleMisc2 As Single
    
    
    ' 2D-Geometery to define the Object's Shape
    Vertex() As mdrVector3  ' Original Vertices (these never change once defined)
    TVertex() As mdrVector3 ' Transformed Vertices (these change all the time)
    Face() As Variant       ' Connect the dots [Vertices] together to form shapes.
    
    ' 2D World Coordinates (ie. The object's position in the game/world.)
    WorldPos As mdrVector3
    
    ' "Vector" stores the object direction and magnitude vector.
    ' Adjust this to tell the object in which direction to travel, and how fast.
    Vector As mdrVector3    ' Direction/Speed Vector (Typically changes when the user presses the arrow keys to move something)
    TVector As mdrVector3   ' Transformed Vector
    
    SpinVector As Single    ' I've used the word 'vector' here to mean a '1 dimensional vector'. Perhaps I should change it to SpinScalar?
    RotationAboutZ As Single
        
    ' Defines the min/max and avg size of an Object.
    ' Useful for detecting collisions against the irregular shape of Asteroids.
    MinSize As Single ' These Min and Max values are helpful in determining how close
    MaxSize As Single ' a player ship can come to an Asteroid, before being damaged or destroyed.
    AvgSize As Single
    
    ' Colour of the object.
    Red As Integer      ' Any integer between 0-255
    Green As Integer    ' Any integer between 0-255
    Blue As Integer     ' Any integer between 0-255
    
    ' Attack/Defence/Strength etc.
    Health As Single    ' Typically between 0 and 100.
    
    
    PreviousThreatLevel As Single
    CurrentThreatLevel As Single

End Type

