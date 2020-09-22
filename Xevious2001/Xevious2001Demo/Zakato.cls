VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Zakato"
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
Dim Direction As Single
Dim Visible As Boolean
Dim VisiblePos As Single

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    With SpritePosition
        SpritePosition.setRelativePosition SpritePosition.RelativeX + Direction, SpritePosition.RelativeY + Abs(geGameEngine.ScrollDownSpeed) + 2, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If SpritePosition.RelativeY > VisiblePos And Visible = False Then
        SpriteRegion.SetRect 4, 4, 24, 24
        SpriteImage = 139
        Visible = True
        T = 0
    End If
    If Visible And SpriteImage > 137 And T > 2 Then
        SpriteImage = SpriteImage - 1
        T = 0
    End If
    T = T + 1
    If SpriteImage = 137 Then
        If geGameEngine.GetRandomNumber() > 0.95 Then
            Dim NewSprite As New Spario
            Set NewSprite.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add NewSprite
            NewSprite.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            NewSprite.PrepareToStart
            
            Dim AirEx As New AirExplosion
            Set AirEx.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add AirEx
            AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            geGameEngine.collDeletedSprites.Add Index
            
        End If
    End If
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "AirWeapon" Then
    
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        
        geGameEngine.Score = geGameEngine.Score + 300

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
    Direction = Sgn(geGameEngine.MousePos.RelativeX - SpritePosition.RelativeX) * geGameEngine.GetRandomNumber
    VisiblePos = geGameEngine.ScrollingRegion.Height * geGameEngine.GetRandomNumber

End Sub



Private Sub Class_Initialize()
    SpriteName = "Zakato"
    SpriteImage = 0
    'SpriteRegion.SetRect 0, 0, 32, 32
    SpriteRegion.SetRect 0, 0, 0, 0
    ZOrder = 60
    Visible = False
End Sub


