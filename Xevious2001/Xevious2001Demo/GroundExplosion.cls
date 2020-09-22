VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroundExplosion"
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
    Static T As Integer
    With SpritePosition
        SpritePosition.setAbsolutePosition SpritePosition.AbsoluteX, SpritePosition.AbsoluteY, geGameEngine.ScrollingRegion
        If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
            geGameEngine.collDeletedSprites.Add Index
        End If
    End With
    T = T + 1
    If T > 3 Then
        SpriteImage = SpriteImage + 1
        If SpriteImage > 116 Then SpriteImage = 110
        T = 0
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

End Sub


Private Sub Class_Initialize()
    SpriteName = "GroundExplosion"
    ZOrder = 0
    SpriteImage = 110
    SpriteRegion.SetRect
End Sub
