VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bakura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SpriteName As String
Public SpritePosition As New Position
Public SpriteRegion As New RectRegion
Public geGameEngine As ctlGameEngine
Public ZOrder As Integer
Dim SpriteImage As Integer
Dim Direction As Integer

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    With SpritePosition
        SpritePosition.setRelativePosition SpritePosition.RelativeX, SpritePosition.RelativeY + Abs(geGameEngine.ScrollDownSpeed) + 1, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    T = T + 1
    If T > 3 Then
        SpriteImage = SpriteImage + 2
        If SpriteImage > 79 And SpriteImage < 100 Then SpriteImage = 134
        If SpriteImage > 134 Then SpriteImage = 71
        T = 0
    End If
    
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "AirWeapon" Then
    
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(CollidedIndex).SpritePosition, geGameEngine.ScrollingRegion
        
        geGameEngine.collDeletedSprites.Add CollidedIndex
        
        
    End If
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY, SpriteImage + 1
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
    SpriteImage = 71 + Int(geGameEngine.GetRandomNumber() * 4) * 2
End Sub



Private Sub Class_Initialize()
    Randomize
    SpriteName = "Bakura"
    SpriteRegion.SetRect 0, 0, 64, 32
    ZOrder = 52
End Sub


