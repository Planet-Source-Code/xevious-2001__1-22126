VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sol"
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
    If Killed Then
        If T > 9 Then
            SpriteImage = SpriteImage + 1
            T = 0
        End If
        If SpriteImage > 36 Then SpriteImage = 36
    End If
    T = T + 1
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
        If SpriteImage = EmptyImage Then
    
            Dim Ex As New AirExplosion
            Set Ex.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add Ex
            Ex.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            
            geGameEngine.Score = geGameEngine.Score + 2000
            
            geGameEngine.collDeletedSprites.Add CollidedIndex
            Killed = True
            SpriteImage = 32
        ElseIf SpriteImage = 36 Then
            Dim Ex2 As New GroundExplosion
            Set Ex2.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add Ex2
            Ex2.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            
            geGameEngine.Score = geGameEngine.Score + 2000
            
            geGameEngine.collDeletedSprites.Add Index
            geGameEngine.collDeletedSprites.Add CollidedIndex
        
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
    SpriteName = "Sol"
    EmptyImage = 0
    SpriteImage = EmptyImage
    SpriteRegion.SetRect 0, 0, 32, 32
    ZOrder = 12
End Sub




