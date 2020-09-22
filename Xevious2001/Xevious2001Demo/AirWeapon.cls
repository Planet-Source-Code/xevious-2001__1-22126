VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirWeapon"
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

Public Sub Move(ByVal Index As Integer)
    With SpritePosition
        .setRelativePosition .RelativeX, .RelativeY - 10, geGameEngine.ScrollingRegion
        If .RelativeY < 0 Then
            geGameEngine.collDeletedSprites.Add Index
        End If
    End With
    If SpriteImage = 129 Then
        SpriteImage = 130
    Else
        SpriteImage = 129
    End If
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If

End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)

End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
End Sub
Public Sub Animation()

End Sub

Public Sub PrepareToStart()
    If geGameEngine.UserAirMissile > 2 Then SpritePosition.setRelativePosition -100, -100, geGameEngine.ScrollingRegion

    geGameEngine.UserAirMissile = geGameEngine.UserAirMissile + 1

End Sub

Private Sub Class_Initialize()
    SpriteName = "AirWeapon"
    ZOrder = 53
    SpriteRegion.SetRect 7, 10, 17, 5
    SpriteImage = 129
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    geGameEngine.UserAirMissile = geGameEngine.UserAirMissile - 1

End Sub
