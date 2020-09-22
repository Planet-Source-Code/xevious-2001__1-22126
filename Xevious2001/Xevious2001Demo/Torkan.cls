VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Torkan"
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
Dim LanchPos As Integer
Dim DirectionX As Integer
Dim DirectionY As Integer

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    With SpritePosition
        SpritePosition.setRelativePosition SpritePosition.RelativeX, SpritePosition.RelativeY + Abs(geGameEngine.ScrollDownSpeed) + DirectionY, geGameEngine.ScrollingRegion
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If Not geGameEngine.IsSpriteInScrollingRegion(Me) Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If LanchPos = 0 Then
        If T > 2 Then
            SpriteImage = SpriteImage + 1
            If SpriteImage > 41 Then
                SpriteImage = 41
                Dim NewSprite As New Spario
                Set NewSprite.geGameEngine = geGameEngine
                geGameEngine.collEnemy.Add NewSprite
                NewSprite.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
                NewSprite.PrepareToStart
                DirectionY = -Abs(geGameEngine.ScrollDownSpeed) - 5
                LanchPos = geGameEngine.ScrollingRegion.Height
            End If
            T = 0
        End If
        T = T + 1
    End If
    If SpritePosition.RelativeY > LanchPos Then
        LanchPos = 0
        DirectionY = 0
    End If
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "AirWeapon" Then
    
        Dim AirEx As New AirExplosion
        Set AirEx.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add AirEx
        AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.Score = geGameEngine.Score + 50
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
    LanchPos = (geGameEngine.ScrollingRegion.Height) * geGameEngine.GetRandomNumber()
    DirectionY = 2
End Sub



Private Sub Class_Initialize()
    SpriteName = "Torkan"
    SpriteImage = 37
    SpriteRegion.SetRect 3, 2, 26, 28
    ZOrder = 60
End Sub



