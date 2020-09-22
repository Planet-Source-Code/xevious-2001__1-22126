VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddorGuileness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SpriteName As String
Public SpritePosition As New Position
Public SpriteRegion As New RectRegion
Public geGameEngine As ctlGameEngine
Public ZOrder As Integer
Dim SpriteImage As Integer
Dim Direction As Single
Dim NoFire As Boolean
Dim OriginalScrollSpeed As Single
Dim Killed As Boolean

Public Sub Move(ByVal Index As Integer)
    Static T As Integer
    Static Firing As Boolean
    With SpritePosition
        SpritePosition.setRelativePosition .RelativeX, .RelativeY - geGameEngine.ScrollDownSpeed + Direction, geGameEngine.ScrollingRegion
        If .RelativeY + 165 > geGameEngine.ScrollingRegion.Height / 2 And Not Killed Then
            Direction = geGameEngine.ScrollDownSpeed
        End If
    End With
    'geGameEngine.pastesprite SpritePosition.relativeX, SpritePosition.relativeY, SpriteImage
    If geGameEngine.ScrollingRegion.Height < SpritePosition.RelativeY Then
        geGameEngine.collDeletedSprites.Add Index
    End If
    If geGameEngine.ScrollingRegion.Top < 500 Then
        geGameEngine.ScrollDownSpeed = geGameEngine.ScrollDownSpeed / 2
        Direction = geGameEngine.ScrollDownSpeed
    End If
    If geGameEngine.ScrollingRegion.Top < 10 Then
        'geGameEngine.ScrollDownSpeed = geGameEngine.ScrollDownSpeed / 2
        Dim NewSprite2 As New Spario
        Dim Newpos2 As New Position
        Set NewSprite2.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add NewSprite2
        Newpos2.setRelativePosition geGameEngine.UserSprite.SpritePosition.RelativeX, geGameEngine.UserSprite.SpritePosition.RelativeY, geGameEngine.ScrollingRegion
        NewSprite2.SpritePosition.CopyPosition Newpos2, geGameEngine.ScrollingRegion
        NewSprite2.PrepareToStart
    End If
    If T > 8 Then
        If geGameEngine.GetRandomNumber() > 0.7 And Not NoFire Then
            Dim NewSprite As New Spario
            Dim Newpos As New Position
            Set NewSprite.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add NewSprite
            Newpos.setRelativePosition geGameEngine.collEnemy(Index).SpritePosition.RelativeX + 64, geGameEngine.collEnemy(Index).SpritePosition.RelativeY + 64, geGameEngine.ScrollingRegion
            NewSprite.SpritePosition.CopyPosition Newpos, geGameEngine.ScrollingRegion
            NewSprite.PrepareToStart
            T = 0
        End If
    End If
    T = T + 1
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).SpriteName = "Weapon" Then
    
        Dim Ex As New AirExplosion
        Set Ex.geGameEngine = geGameEngine
        geGameEngine.collEnemy.Add Ex
        Ex.SpritePosition.CopyPosition geGameEngine.collEnemy(CollidedIndex).SpritePosition, geGameEngine.ScrollingRegion
        geGameEngine.Score = geGameEngine.Score + 4000
        OriginalScrollSpeed = geGameEngine.ScrollDownSpeed

        'geGameEngine.collDeletedSprites.Add Index
        geGameEngine.collDeletedSprites.Add CollidedIndex
        SpriteImage = 85
        Direction = 1
        Killed = True
        NoFire = True
        
    End If
End Sub
Public Sub Show()
    With SpritePosition
        For x = 0 To 4
            For y = 0 To 4
                geGameEngine.PasteSprite .RelativeX + 32 * x, .RelativeY + 32 * y, 140 + x + y * 5
            Next y
        Next x
    End With
End Sub
Public Sub Animation(ByVal Index As Integer)

End Sub

Public Sub PrepareToStart()
    OriginalScrollSpeed = geGameEngine.ScrollDownSpeed
End Sub



Private Sub Class_Initialize()
    SpriteName = "AddorGuileness"
    SpriteImage = 133
    SpriteRegion.SetRect 64, 64, 32, 32
    ZOrder = 10
    Direction = 0.5
End Sub






Private Sub Class_Terminate()
    geGameEngine.ScrollDownSpeed = OriginalScrollSpeed

End Sub
