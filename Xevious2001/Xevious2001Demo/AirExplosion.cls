VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirExplosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Implements ISprite
Public SpriteName As String
Public SpritePosition As New Position
Public SpriteRegion As New RectRegion
Public geGameEngine As ctlGameEngine
Public ZOrder As Integer
Dim SpriteImage As Integer

Public Sub Move(ByVal Index As Integer)
    With SpritePosition
        SpritePosition.setAbsolutePosition SpritePosition.AbsoluteX, SpritePosition.AbsoluteY, geGameEngine.ScrollingRegion
        If SpriteImage = 0 Then
            'geGameEngine.RemoveSprite (Index)
            geGameEngine.collDeletedSprites.Add Index
        End If
    End With

End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)

End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
    SpriteImage = SpriteImage + 1
    If SpriteImage > 109 Then SpriteImage = 0
End Sub
Public Sub Animation()

End Sub

Public Sub PrepareToStart()

End Sub


Private Sub Class_Initialize()
    SpriteName = "AirExplosion"
    ZOrder = 60
    SpriteImage = 99
    SpriteRegion.SetRect 0, 0, 0, 0
End Sub

Private Sub Class_Terminate()
        'MsgBox "end!"

End Sub

