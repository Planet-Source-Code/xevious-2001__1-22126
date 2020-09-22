VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GaruDerota"
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
Dim NoFire As Boolean

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    Static Firing As Boolean
    With SpritePosition
        SpritePosition.setAbsolutePosition SpritePosition.AbsoluteX, SpritePosition.AbsoluteY, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    'If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
    If geGameEngine.ScrollingRegion.Height < SpritePosition.RelativeY Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If T > 5 Then
        If geGameEngine.GetRandomNumber() > 0.8 And Not NoFire Then
            Dim NewSprite As New Spario
            Dim Newpos As New Position
            Set NewSprite.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add NewSprite
            Newpos.setRelativePosition geGameEngine.collEnemy(Index).SpritePosition.RelativeX + 32, geGameEngine.collEnemy(Index).SpritePosition.RelativeY + 32, geGameEngine.ScrollingRegion
            NewSprite.SpritePosition.CopyPosition Newpos, geGameEngine.ScrollingRegion
            NewSprite.PrepareToStart
            T = 0
        End If
    End If
    T = T + 1
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
    
        Dim Ex As New AirExplosion
        Set Ex.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add Ex
        Ex.SpritePosition.CopyPosition geGameEngine.collEnemy(CollidedIndex).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.Score = geGameEngine.Score + 2000
        
        'geGameEngine.collDeletedSprites.Add Index
        geGameEngine.collDeletedSprites.Add CollidedIndex
        SpriteImage = 85
        NoFire = True
        
    End If
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, 81
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY, 82
    geGameEngine.PasteSprite SpritePosition.RelativeX + 64, SpritePosition.RelativeY, 83
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY + 32, 84
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY + 32, SpriteImage
    geGameEngine.PasteSprite SpritePosition.RelativeX + 64, SpritePosition.RelativeY + 32, 86
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY + 64, 87
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY + 64, 88
    geGameEngine.PasteSprite SpritePosition.RelativeX + 64, SpritePosition.RelativeY + 64, 89
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
End Sub



Private Sub Class_Initialize()
    SpriteName = "GaruDerota"
    SpriteImage = 133
    SpriteRegion.SetRect 32, 32, 32, 32
    ZOrder = 10
    Randomize
End Sub





