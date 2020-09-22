VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Discobolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SpriteName As String
Public SpritePosition As New Position
Public SpriteRegion As New RectRegion
Public geGameEngine As GameEngine
Public ZOrder As Integer
Dim SpriteImage As Integer
Dim Direction As Integer

Public Sub Move(ByVal Index As Integer)
    With SpritePosition
        SpritePosition.setrelativePosition SpritePosition.relativeX + Direction, SpritePosition.relativeY + 2, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    SpriteImage = SpriteImage + 1
    If SpriteImage > 52 Then SpriteImage = 47
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "AirWeapon" Then
    
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        
        geGameEngine.collDeletedSprites.Add Index
        geGameEngine.collDeletedSprites.Add CollidedIndex
        
        'If Index > CollidedIndex Then
        '    geGameEngine.collEnemy(Index).Remove
        '    geGameEngine.collEnemy(CollidedIndex).Remove
        'Else
        '    geGameEngine.collEnemy.Remove (CollidedIndex)
        '    geGameEngine.collEnemy.Remove (Index)
        'End If
        
    End If
End Sub
Public Sub Show()
    geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
    Direction = Sgn(geGameEngine.MousePos.relativeX - SpritePosition.relativeX)
End Sub



Private Sub Class_Initialize()
    SpriteName = "Discobolo"
    SpriteImage = 47
    SpriteRegion.SetRect 0, 0, 32, 32
    ZOrder = 60
End Sub

