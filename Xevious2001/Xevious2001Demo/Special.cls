VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Special"
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
Dim Killed As Boolean
Dim EmptyImage As Integer

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    With SpritePosition
        SpritePosition.setAbsolutePosition SpritePosition.AbsoluteX, SpritePosition.AbsoluteY, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    T = T + 1
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
        If SpriteImage = EmptyImage Then
    
            Killed = True
            SpriteImage = 31
            ZOrder = 80
        ElseIf SpriteImage = 31 Then
            
            'If geGameEngine.collEnemy(CollidedIndex).SpriteName = geGameEngine.UserSpriteName Then
            '    frmGame.NumSolvalou = frmGame.NumSolvalou + 1
            '    Dim Ex As New AirExplosion
            '    Set Ex.geGameEngine = geGameEngine
            '    geGameEngine.collEnemy.Add Ex
            '    Ex.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            '
            '    geGameEngine.collDeletedSprites.Add Index
            'End If
        End If
    End If
End Sub
Public Sub Show()
    geGameEngine.PasteSprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
End Sub



Private Sub Class_Initialize()
    SpriteName = "Special"
    EmptyImage = 0
    SpriteImage = EmptyImage
    SpriteRegion.SetRect 0, 0, 32, 32
    ZOrder = 12
End Sub





