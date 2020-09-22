VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Garubara"
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
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
    
        Dim Ex As New AirExplosion
        Set Ex.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add Ex
        Ex.SpritePosition.CopyPosition geGameEngine.collEnemy(CollidedIndex).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.Score = geGameEngine.Score + 300
        
        'geGameEngine.collDeletedSprites.Add Index
        geGameEngine.collDeletedSprites.Add CollidedIndex
        SpriteImage = 94
        NoFire = True
        
    End If
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, 90
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY, 91
    geGameEngine.PasteSprite SpritePosition.RelativeX + 64, SpritePosition.RelativeY, 92
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY + 32, 93
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY + 32, SpriteImage
    geGameEngine.PasteSprite SpritePosition.RelativeX + 64, SpritePosition.RelativeY + 32, 95
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY + 64, 96
    geGameEngine.PasteSprite SpritePosition.RelativeX + 32, SpritePosition.RelativeY + 64, 97
    geGameEngine.PasteSprite SpritePosition.RelativeX + 64, SpritePosition.RelativeY + 64, 98
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
End Sub



Private Sub Class_Initialize()
    SpriteName = "Garubara"
    SpriteImage = 136
    SpriteRegion.SetRect 32, 32, 32, 32
    ZOrder = 10
    Randomize
End Sub




