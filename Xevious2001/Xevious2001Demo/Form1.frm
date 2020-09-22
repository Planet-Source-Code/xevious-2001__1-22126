VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xevious2000"
   ClientHeight    =   6975
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   StartUpPosition =   3  'Windows Default
   Begin Xevious2000.ctlGameEngine GameEngine1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      _extentx        =   8493
      _extenty        =   12303
   End
   Begin VB.Timer Timer1 
      Left            =   7440
      Top             =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Game 
      Caption         =   "Game"
      Begin VB.Menu Start 
         Caption         =   "Start"
         Shortcut        =   {F2}
      End
      Begin VB.Menu PauseGame 
         Caption         =   "Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu RestartGame 
         Caption         =   "Restart"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Option 
      Caption         =   "Option"
      Begin VB.Menu CaptureMouse 
         Caption         =   "CaptureMouseInForm"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShipScore As Long
Dim Level As Integer
Public NumSolvalou As Integer
Dim Killed As Integer
Dim DemoMode As Boolean
Dim TimerInterval As Integer

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Private Sub DisableTrap(CurForm As Form)
    Dim erg As Long
    Dim NewRect As RECT
    With NewRect
      .Left = 0&
      .Top = 0&
      .Right = Screen.Width / Screen.TwipsPerPixelX
      .Bottom = Screen.Height / Screen.TwipsPerPixelY
    End With
    erg& = ClipCursor(NewRect)
End Sub

Private Sub EnableTrap(CurForm As Form)
    Dim x As Long, y As Long, erg As Long
    Dim NewRect As RECT
    x& = Screen.TwipsPerPixelX
    y& = Screen.TwipsPerPixelY
    'of the form
    With NewRect
    .Left = CurForm.Left / x&
    .Top = CurForm.Top / y&
    .Right = .Left + CurForm.Width / x&
    .Bottom = .Top + CurForm.Height / y&
    End With
    erg& = ClipCursor(NewRect)
End Sub


Private Sub Command2_Click()
    Dim pippo2 As New GroundExplosion
    Set pippo2.geGameEngine = GameEngine1
    GameEngine1.CreateNewSprite pippo2
    pippo2.SpritePosition.setRelativePosition 10, 10, GameEngine1.ScrollingRegion
    
    Dim pippo3 As New AirExplosion
    Set pippo3.geGameEngine = GameEngine1
    GameEngine1.CreateNewSprite pippo3
    pippo3.SpritePosition.setRelativePosition 50, 50, GameEngine1.ScrollingRegion

End Sub

Private Sub Command3_Click()
    Dim pippo2 As New Discobolo
    Set pippo2.geGameEngine = GameEngine1
    GameEngine1.CreateNewSprite pippo2
    pippo2.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
    pippo2.PrepareToStart
    
End Sub

Private Sub Command4_Click()

End Sub

Private Sub CaptureMouse_Click()
    If Not CaptureMouse.Checked Then
        EnableTrap Me
    End If
    CaptureMouse.Checked = Not CaptureMouse.Checked
End Sub

Private Sub Exit_Click()
    EndGame
    DisableTrap Me
    
    End
End Sub

Private Sub Form_Load()
    Randomize
    Form2.Show
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    DoEvents
    GameEngine1.LoadGraphics App.Path + "\backgroundbn.bmp", App.Path + "\sprites.bmp"
    
    GameEngine1.ScoreLabel.FontBold = True
    GameEngine1.ScoreLabel.FontSize = 14
    GameEngine1.ScoreLabel.ForeColor = RGB(255, 255, 255)
    
    'GameEngine1.LoadGraphics "", App.Path + "\sprites.bmp"
    'GameEngine1.DimensionRatio = 1
    'GameEngine1.BackgroundDimensionRatio = 1
    'GameEngine1.SpriteDimensionRatio = 1.5
    GameEngine1.ScrollDownSpeed = -2
    GameEngine1.LoadHighScore App.Path + "\Xevious.hsc"
    GameEngine1.Picture.MousePointer = 2
    Level = 1
    SetLevelStartigPos
    TimerInterval = 80
    StartDemoMode
    
    Unload Form2

End Sub


Private Sub Form_Unload(Cancel As Integer)
    EndGame
    DisableTrap Me
End Sub

Private Sub GameEngine1_EndScrollingRegion(Direction As dnzDirection)
    'New Level
    NewLevel
End Sub

Private Sub GameEngine1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GameEngine1.UserFire (Button)
End Sub

Private Sub GameEngine1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    GameEngine1.MousePos.setRelativePosition x - 16, y - 16, GameEngine1.ScrollingRegion
    
End Sub

Private Sub GameEngine1_NewMapObject(ObjectMapParam As String)
    If Killed Then Exit Sub
    Dim obj As Object
    'Exit Sub
    Select Case LCase(Split(ObjectMapParam, ",")(2))
        Case "toroid"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Toroid
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "jara"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Jara
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "zoshi"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Zoshi
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "torkan"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Torkan
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "kapi"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Kapi
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "terrazi"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Terrazi
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "zakato"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Zakato
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i


        Case "bakura"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Bakura
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition GameEngine1.ScrollingRegion.Width * GameEngine1.GetRandomNumber(), -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "zolbak"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Zolbak
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "logram"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Logram
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + (Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        

        Case "domogram"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Domogram
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "barra"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Barra
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "garubara"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Garubara
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 150 * i, -96, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "garuderota"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New GaruDerota
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 150 * i, -96, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "derota"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Derota
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "grobda"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Grobda
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "sol"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Sol
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

        Case "special"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New Special
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 40 * i, -32, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i
        Case "addorguileness"
            For i = 0 To Val(UCase(Split(ObjectMapParam, ",")(3))) - 1
                Set obj = New AddorGuileness
                Set obj.geGameEngine = GameEngine1
                GameEngine1.CreateNewSprite obj
                obj.SpritePosition.setRelativePosition -GameEngine1.ScrollingRegion.Left + Val(Split(ObjectMapParam, ",")(0)) + 150 * i, -160, GameEngine1.ScrollingRegion
                obj.PrepareToStart
            Next i

    End Select
    
End Sub

Private Sub GameEngine1_UserCollision()
    If NumSolvalou = 1 Then
        GameEngine1.SetBackGroundText "GAME OVER", 60
        Killed = True
    Else
        Killed = True
    End If
End Sub

Private Sub PauseGame_Click()
    Timer1.Interval = 0
End Sub

Private Sub RestartGame_Click()
    Timer1.Interval = TimerInterval

End Sub

Private Sub Start_Click()
    StartGame
End Sub

Private Sub Timer1_Timer()
    Static T As Long
    If Not DemoMode Then
        If Killed = False Then
            GameEngine1.TimeClick
            GameEngine1.ScoreLabel.Caption = Format(GameEngine1.Score, "000000") + "                                      " + Format(NumSolvalou - 1)
            Label1.Caption = GameEngine1.collEnemy.Count
        Else
            T = T + 1
            GameEngine1.TimeClick
            If T > 60 Then
                NumSolvalou = NumSolvalou - 1
                If NumSolvalou < 1 Then
                    EndGame
                Else
                    GameEngine1.SetBackGroundText "", 0
                    StartLevel
                    Killed = False
                    T = 0
                End If
            End If
        End If
    Else
        GameEngine1.TimeClick
    End If
    If GameEngine1.Score > ShipScore Then
        NumSolvalou = NumSolvalou + 1
        ShipScore = ShipScore + 60000
    End If
End Sub


Sub StartGame()
    DemoMode = False
    Level = 1
    ShipScore = 20000
    NumSolvalou = 3
    StartLevel
    GameEngine1.Score = 0
    'GameEngine1.ScollTextLabel.Top = (GameEngine1.Height - GameEngine1.ScollTextLabel.Height) / 2

End Sub
Sub EndGame()
    If GameEngine1.IsInHighScore(GameEngine1.Score) Then
        a$ = InputBox("Insert your name for High-score")
        GameEngine1.UpdateHighScore a$, GameEngine1.Score
        GameEngine1.SaveHighScore App.Path + "\Xevious.hsc"
        GameEngine1.Score = 0
    End If
    GameEngine1.ClearAllSprite
    StartDemoMode
End Sub
Sub SetLevelStartigPos()
            startx = 0
    
    'GameEngine1.SetStartingPosition startx, GameEngine1.BackgroundImageDimension.Height - GameEngine1.Height * GameEngine1.DimensionRatio
    GameEngine1.SetStartingPosition startx, GameEngine1.BackgroundImageDimension.Height - GameEngine1.Height

End Sub
Sub CreateSolvalou()
    Dim lSolvalou As New Solvalou
    Set lSolvalou.geGameEngine = GameEngine1
    GameEngine1.CreateNewSprite lSolvalou, "Solvalou"
    GameEngine1.UserSpriteName = "Solvalou"
    lSolvalou.PrepareToStart
    Set GameEngine1.UserSprite = lSolvalou

End Sub
Sub NewLevel()
    'GameEngine1.SetBackGroundText "Completed level " + Format(Level), 40
    Level = Level + 1
    SetLevelStartigPos
    StartLevel
End Sub
Sub StartLevel()
    
    GameEngine1.ClearAllSprite
    GameEngine1.TimeClick
    SetLevelStartigPos
    Killed = False

    GameEngine1.LoadMap App.Path + "\level1" + ".map"
    'GameEngine1.LoadMap App.Path + "\map.txt"
    GameEngine1.TimeClick
    Timer1.Interval = TimerInterval
    If Not DemoMode Then
        CreateSolvalou
        GameEngine1.SetBackGroundText "Starting level " + Format(Level), 40
    Else
        GameEngine1.ShowHighScore
    End If
End Sub
Sub StartDemoMode()
    DemoMode = True
    Level = 1
    NumSolvalou = 3
    StartLevel
End Sub
