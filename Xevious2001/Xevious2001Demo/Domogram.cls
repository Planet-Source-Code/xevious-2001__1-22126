VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Domogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SpriteName As String
Public SpritePosition As New Position
Public SpriteRegion As New RectRegion
Public geGameEngine As ctlGameEngine
Public ZOrder As Integer
Public PathType As Integer
Dim SpriteImage As Integer
Dim Direction As Integer
Dim Firing As Boolean

Public Sub Move(ByVal Index As Integer)
    Dim IncX As Single
    Dim IncY As Single
    Static T As Integer
    Static t1 As Single
    geGameEngine.SpritePathMove t1, PathType, IncX, IncY
    With SpritePosition
        SpritePosition.setAbsolutePosition SpritePosition.AbsoluteX + IncX, SpritePosition.AbsoluteY + IncY, geGameEngine.ScrollingRegion
    End With
    t1 = t1 + 1
    If geGameEngine.ScrollingRegion.Height < SpritePosition.RelativeY Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If geGameEngine.GetRandomNumber() > 0.99 Then
        Firing = True
    End If
    If Firing Then
        T = T + 1
        If T > 3 Then
            SpriteImage = SpriteImage + 1
            If SpriteImage > 6 Then
                Dim NewSprite As New Spario
                Set NewSprite.geGameEngine = geGameEngine
                geGameEngine.collEnemy.Add NewSprite
                NewSprite.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
                NewSprite.PrepareToStart
                SpriteImage = 3
                Firing = False
            End If
            T = 0
        End If
    End If
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
    
        Dim AirEx As New GroundExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        
        geGameEngine.Score = geGameEngine.Score + 800

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
    PathType = Int(geGameEngine.GetRandomNumber() * 6)
End Sub



Private Sub Class_Initialize()
    Randomize
    SpriteName = "Domogram"
    SpriteImage = 3
    SpriteRegion.SetRect 0, 0, 32, 32
    ZOrder = 10
End Sub


