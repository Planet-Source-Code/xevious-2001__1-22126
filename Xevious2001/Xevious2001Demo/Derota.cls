VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Derota"
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

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    Static Firing As Boolean
    With SpritePosition
        SpritePosition.setAbsolutePosition SpritePosition.AbsoluteX, SpritePosition.AbsoluteY, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If T > 10 Then
        If geGameEngine.GetRandomNumber() > 0.95 Then
            Dim NewSprite As New Spario
            Set NewSprite.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add NewSprite
            NewSprite.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            NewSprite.PrepareToStart
            T = 0
        End If
    End If
    T = T + 1
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
    
        Dim GrEx As New GroundExplosion
        Set GrEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add GrEx
        GrEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.Score = geGameEngine.Score + 1000

        
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
End Sub



Private Sub Class_Initialize()
    SpriteName = "Derota"
    SpriteImage = 2
    SpriteRegion.SetRect 3, 3, 26, 26
    ZOrder = 10
    Randomize
End Sub




