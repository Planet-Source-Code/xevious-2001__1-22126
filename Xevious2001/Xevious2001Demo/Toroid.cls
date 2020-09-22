VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Toroid"
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
Dim Direction As Single

Public Sub Move(ByVal Index As Integer)
    Static Rolling As Boolean
    
    With SpritePosition
        SpritePosition.setRelativePosition SpritePosition.RelativeX + Direction, SpritePosition.RelativeY + Abs(geGameEngine.ScrollDownSpeed) + 1, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If Not Rolling And Abs(SpritePosition.RelativeX - geGameEngine.UserSpritePosition.RelativeX) < 4 Then
        Rolling = True
        Direction = -Direction
    End If
    If Rolling Then
        SpriteImage = SpriteImage + 1
        Direction = Direction + Sgn(Direction) * 0.5
        If Abs(Direction) > 4 Then Direction = 4 * Sgn(Direction)
        If SpriteImage > 52 Then SpriteImage = 47
    End If
    If geGameEngine.GetRandomNumber() > 0.995 Then
        Dim NewSprite As New Spario
        Set NewSprite.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add NewSprite
        NewSprite.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        NewSprite.PrepareToStart
    End If
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "AirWeapon" Then
    
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        
        geGameEngine.Score = geGameEngine.Score + 30

        geGameEngine.collDeletedSprites.Add Index
        geGameEngine.collDeletedSprites.Add CollidedIndex
        
    End If
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
    Direction = Sgn(geGameEngine.MousePos.RelativeX - SpritePosition.RelativeX)
End Sub



Private Sub Class_Initialize()
    SpriteName = "Toroid"
    SpriteImage = 47
    SpriteRegion.SetRect 0, 0, 32, 32
    ZOrder = 60
End Sub

