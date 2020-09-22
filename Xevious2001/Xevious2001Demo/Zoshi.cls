VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Zoshi"
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
Dim LanchPos As Integer
Dim DirectionX As Single
Dim DirectionY As Single
Dim Axes As Single

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    With SpritePosition
        SpritePosition.setRelativePosition SpritePosition.RelativeX + DirectionX * 2, SpritePosition.RelativeY + Abs(geGameEngine.ScrollDownSpeed) + DirectionY, geGameEngine.ScrollingRegion
    End With
    If SpritePosition.RelativeY > geGameEngine.ScrollingRegion.Height Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    Centerscreenx = (Axes / 2 - SpritePosition.RelativeX) / geGameEngine.ScrollingRegion.Width
    If Sgn(DirectionX) <> Sgn(DirectionX + Centerscreenx * 2) Then
        T = T + 1
        If T > 3 Then
            Dim NewSprite As New Spario
            Set NewSprite.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add NewSprite
            NewSprite.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            NewSprite.PrepareToStart
            T = 0
        End If
    End If
    DirectionX = DirectionX + Centerscreenx * 2
    SpriteImage = SpriteImage + 1
    If SpriteImage > 56 Then SpriteImage = 54

End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "AirWeapon" Then
    
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.Score = geGameEngine.Score + 70
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
    DirectionY = 1.5
    Axes = geGameEngine.GetRandomNumber() * geGameEngine.ScrollingRegion.Width
End Sub



Private Sub Class_Initialize()
    SpriteName = "Zoshi"
    SpriteImage = 54
    SpriteRegion.SetRect 1, 1, 30, 30
    ZOrder = 60
End Sub






