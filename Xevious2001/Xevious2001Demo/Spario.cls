VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spario"
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
Dim DirectionX As Single
Dim DirectionY As Single

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    SpritePosition.setRelativePosition SpritePosition.RelativeX + DirectionX, SpritePosition.RelativeY + DirectionY, geGameEngine.ScrollingRegion
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    OkDirX = -(SpritePosition.RelativeX - geGameEngine.UserSpritePosition.RelativeX)
    OkDirY = -(SpritePosition.RelativeY - geGameEngine.UserSpritePosition.RelativeY)
    If Sgn(DirectionX) <> Sgn(OkDirX) Then
        If Abs(OkDirX) > 100 Then
            Dim AirEx As New AirExplosion
            Set AirEx.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add AirEx
            AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            geGameEngine.collDeletedSprites.Add Index
        End If
    End If
    If Sgn(DirectionY) <> Sgn(OkDirY) Then
        If Abs(OkDirY) > 100 Then
            Dim AirEx2 As New AirExplosion
            Set AirEx2.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add AirEx2
            AirEx2.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            geGameEngine.collDeletedSprites.Add Index
        End If
    End If
    
    T = T + 1
    If T > 4 Then
        T = 0
        If SpriteImage = 132 Then
            SpriteImage = 131
        Else
            SpriteImage = 132
        End If
    End If
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
    Dx = (geGameEngine.UserSpritePosition.RelativeX + 16 - SpritePosition.RelativeX) + 0.0001
    Dy = (geGameEngine.UserSpritePosition.RelativeY + 16 - SpritePosition.RelativeY) + 0.0001
    s = (Dx ^ 2 + Dy ^ 2) ^ 0.5
    DirectionX = 4 * Dx / s
    DirectionY = 4 * Dy / s
    
End Sub



Private Sub Class_Initialize()
    SpriteName = "Spario"
    SpriteImage = 131
    SpriteRegion.SetRect 12, 12, 7, 7
    ZOrder = 52
End Sub


