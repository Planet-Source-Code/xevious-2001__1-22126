VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Weapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SpriteName As String
Public SpritePosition As New Position
Public SpriteRegion As New RectRegion
Public geGameEngine As ctlGameEngine
Public ZOrder As Integer

Dim Direction As Integer
Public Sub Move(ByVal Index As Integer)
    Static TTL As Integer
    Static T As Integer
    With SpritePosition
        If T > 4 Then
            NewX = 0
        Else
            NewX = 1 * Direction
        End If
        NewY = 0.5 * T ^ 2
        If NewY > 12 Then NewY = 12
        .setRelativePosition .RelativeX + NewX, .RelativeY - NewY, geGameEngine.ScrollingRegion
        T = T + 1
        If .RelativeY < 0 Then
            geGameEngine.collDeletedSprites.Add Index
        End If
    End With
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    TTL = TTL + 1
    If TTL > 12 Then
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.collDeletedSprites.Add Index
    End If
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    'If geGameEngine.collEnemy(CollidedIndex).SpriteName <> "Solvalou" Then
    '    If geGameEngine.collEnemy(CollidedIndex).ZOrder > 49 Then
    '
    '    End If
    'End If
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, 128
End Sub
Public Sub Animation()

End Sub

Public Sub PrepareToStart()
    If geGameEngine.UserGroundMissile > 4 Then SpritePosition.setRelativePosition -100, -100, geGameEngine.ScrollingRegion
    geGameEngine.UserGroundMissile = geGameEngine.UserGroundMissile + 1
    Direction = Sgn(geGameEngine.GetRandomNumber() - 0.5)
End Sub

Private Sub Class_Initialize()
    SpriteName = "Weapon"
    ZOrder = 47
    SpriteRegion.SetRect 14, 9, 3, 11
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    geGameEngine.UserGroundMissile = geGameEngine.UserGroundMissile - 1
End Sub
