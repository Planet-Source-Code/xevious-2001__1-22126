VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Solvalu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public SpriteRegion As New RectRegion
Public SpriteName As String
Public SpritePosition As New Position
Public geGameEngine As GameEngine
Public ZOrder As Integer
Dim SpriteImage As Integer

Public Sub Move(ByVal Index As Integer)
    With SpritePosition
        .setRelativePosition .RelativeX + (Abs(geGameEngine.MousePos.RelativeX - .RelativeX) ^ (1 / 3)) * Sgn(geGameEngine.MousePos.RelativeX - .RelativeX), .RelativeY + (Abs(geGameEngine.MousePos.RelativeY - .RelativeY) ^ (1 / 3)) * Sgn(geGameEngine.MousePos.RelativeY - .RelativeY), geGameEngine.ScrollingRegion
        If geGameEngine.MousePos.RelativeX - .RelativeX < 0 Then
            'Verso sinistra
            SpriteImage = 64
        Else
            SpriteImage = 65
        End If
        If Abs(geGameEngine.MousePos.RelativeX - .RelativeX) < 1 Then
            .setRelativePosition geGameEngine.MousePos.RelativeX, .RelativeY, geGameEngine.ScrollingRegion
            SpriteImage = 63

        End If
    End With

End Sub

Public Sub Show()
    geGameEngine.pastesprite SpritePosition.RelativeX, SpritePosition.RelativeY, SpriteImage
    
End Sub
Public Sub Animation()

End Sub

Public Sub PrepareToStart()
    With geGameEngine.ScrollingRegion
       SpritePosition.setRelativePosition .Width / 2, .Height / 2, geGameEngine.ScrollingRegion
    End With

End Sub
Public Sub Fire(ByVal FireType As Integer)
    If FireType = 2 Then
        Dim pippo As New Weapon
        Set pippo.geGameEngine = geGameEngine
        geGameEngine.CreateNewSprite pippo
        pippo.SpritePosition.CopyPosition SpritePosition, geGameEngine.ScrollingRegion
        pippo.PrepareToStart
    Else
        Dim pippo2 As New AirWeapon
        Set pippo2.geGameEngine = geGameEngine
        geGameEngine.CreateNewSprite pippo2
        pippo2.SpritePosition.CopyPosition SpritePosition, geGameEngine.ScrollingRegion
        pippo2.PrepareToStart
    End If
End Sub

Private Sub Class_Initialize()
    SpriteName = "Solvalou"
    SpriteImage = 63
    ZOrder = 90
    SpriteRegion.SetRect 4, 3, 24, 29
 
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
    If geGameEngine.collEnemy(CollidedIndex).ZOrder < 50 Then Exit Sub
    Select Case geGameEngine.collEnemy(CollidedIndex).SpriteName
        Case "AirWeapon", "Weapon"
        Case Else
            Dim AirEx As New AirExplosion
            Set AirEx.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add AirEx
            AirEx.SpritePosition.CopyPosition geGameEngine.collEnemy(Index).SpritePosition, geGameEngine.ScrollingRegion
            Dim AirEx2 As New AirExplosion
            Set AirEx2.geGameEngine = geGameEngine
            geGameEngine.collEnemy.Add AirEx2
            AirEx2.SpritePosition.CopyPosition geGameEngine.collEnemy(CollidedIndex).SpritePosition, geGameEngine.ScrollingRegion
            
            geGameEngine.collDeletedSprites.Add Index
            geGameEngine.collDeletedSprites.Add CollidedIndex
            
            geGameEngine.UserSpriteCollision
    End Select

End Sub
