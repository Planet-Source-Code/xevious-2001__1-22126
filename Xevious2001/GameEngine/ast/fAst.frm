VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   3225
   ClientTop       =   1920
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8925
   Begin VB.Label Game 
      Caption         =   "Game"
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Sprites:
Private cResShip As cSpriteBitmaps
Private cShip As cSprite
Private cResRockxL As cSpriteBitmaps
Private cResRockL As cSpriteBitmaps
Private cResRockM As cSpriteBitmaps
Private cResRockS As cSpriteBitmaps
Private cRocks() As cSprite

' Background and staging area:
Private cStage As cBitmap
Private cT As cTile

' Get Key presses:
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' General for game:
Private m_bInGame As Boolean


Private Sub GameLoop()
Static bGameLoop As Boolean
Dim i As Long
Dim iSpriteNum As Long
Dim lHDC As Long
Dim lH As Long
Dim lW As Long
Static bLastHyperspace As Boolean

    bGameLoop = Not (bGameLoop)
    iSpriteNum = UBound(cRocks())
    lHDC = Me.hDC
    lW = Me.ScaleWidth \ Screen.TwipsPerPixelX
    lH = Me.ScaleHeight \ Screen.TwipsPerPixelY
    'cShip.Active = True
    cShip.X = (lW - cShip.Width) \ 2
    cShip.Y = (lH - cShip.Height) \ 2
    cShip.Cell = 1
    bLastHyperspace = False
    
    If (bGameLoop) Then
        cStage.RenderBitmap lHDC, 0, 0
    End If
    m_bInGame = bGameLoop
    
    Do While bGameLoop
        
        ' ******************************************************
        ' 1) Firstly, we restore the stage bitmap to its original
        ' state:
        For i = 5 To iSpriteNum
            If (cRocks(i).Active) Then
                cRocks(i).RestoreBackground cStage.hDC
            End If
        Next i
        cShip.RestoreBackground cStage.hDC
        ' ******************************************************
        
        ' (At this point you could modify the background in cStage)
        
        ' ******************************************************
        ' 2) Secondly, we move all the sprites to their new position
        ' on the stage bitmap and copy the stage at that point:
        
        ' Draw rocks:
        For i = 5 To iSpriteNum
        
            If (cRocks(i).Active) Then
                'Debug.Print i, cRocks(i).Width, cRocks(i).Height, cRocks(i).X, cRocks(i).Y, cRocks(i).XDir, cRocks(i).YDir
                ' Determine the new position:
                cRocks(i).IncrementPosition
                
                ' Check for wrap:
                If (cRocks(i).XDir < 0) Then
                    If (cRocks(i).X < -cRocks(i).Width) Then cRocks(i).X = lW + cRocks(i).Width
                Else
                    If (cRocks(i).X > lW) Then cRocks(i).X = -cRocks(i).Width
                End If
                If (cRocks(i).YDir < 0) Then
                    If (cRocks(i).Y < -cRocks(i).Height) Then cRocks(i).Y = lH + cRocks(i).Height
                Else
                    If (cRocks(i).Y > lH) Then cRocks(i).Y = -cRocks(i).Height
                End If
                
                cRocks(i).StoreBackground cStage.hDC, cRocks(i).X, cRocks(i).Y
            End If
        Next i
        ' Spaceship:
        If (GetAsyncKeyState(vbKeyLeft) <> 0) Then
            cShip.Cell = cShip.Cell + 1
            If (cShip.Cell > 12) Then cShip.Cell = 1
        ElseIf (GetAsyncKeyState(vbKeyRight) <> 0) Then
            cShip.Cell = cShip.Cell - 1
            If (cShip.Cell < 1) Then cShip.Cell = 12
        End If
        If (GetAsyncKeyState(vbKeyUp) <> 0) Then
            ' accelerate in current direction:
            
        End If
        If (GetAsyncKeyState(vbKeySpace) <> 0) Then
            If Not bLastHyperspace Then
                ' hyperspace
                bLastHyperspace = True
                cShip.X = Rnd * lW
                cShip.Y = Rnd * lH
            End If
        Else
            bLastHyperspace = False
        End If
        cShip.StoreBackground cStage.hDC, cShip.X, cShip.Y
        
        
        ' ******************************************************
        
        
        ' ******************************************************
        ' 3) Next we draw all the sprites onto the stage:
        For i = 5 To iSpriteNum
            If (cRocks(i).Active) Then
                ' Draw the sprite onto the stage in the new position:
                cRocks(i).TransparentDraw cStage.hDC, cRocks(i).X, cRocks(i).Y, 1, False
            End If
        Next i
        cShip.TransparentDraw cStage.hDC, cShip.X, cShip.Y, cShip.Cell, False

        
        ' ******************************************************

        ' ******************************************************
        
        ' ******************************************************
        ' 3) Finally we transfer the changes in the stage onto
        ' the screen, minimising the number of visible screen
        ' blits as best as we can:
        For i = 5 To iSpriteNum
            If (cRocks(i).Active) Then
                cRocks(i).StageToScreen lHDC, cStage.hDC
            End If
        Next i
        cShip.StageToScreen lHDC, cStage.hDC
        ' ******************************************************
        
        DoEvents
    Loop
        
End Sub


Private Sub CreateSpriteResource( _
        ByRef cR As cSpriteBitmaps, _
        ByVal sFile As String, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal lTransColor As Long _
    )
    Set cR = New cSpriteBitmaps
    cR.CreateFromFile sFile, cX, cY, , lTransColor
End Sub
Private Sub CreateSprite( _
        ByRef cR As cSpriteBitmaps, _
        ByRef cS As cSprite _
    )
    Set cS = New cSprite
    cS.SpriteData = cR
    cS.Create Me.hDC
End Sub
Private Sub InitRockPosition( _
        ByRef cS As cSprite _
    )
    Do While cS.XDir = 0 And cS.YDir = 0
        cS.XDir = ((Rnd * 8) - 4) * 2
        cS.YDir = ((Rnd * 8) - 4) * 2
    Loop
    cS.X = Rnd * (Me.ScaleWidth \ Screen.TwipsPerPixelX \ 2)
    cS.Y = Rnd * (Me.ScaleHeight \ Screen.TwipsPerPixelY \ 2)
    cS.Active = True
    
End Sub
        

Private Sub Form_Load()
Dim i As Integer
Dim lW As Long, lH As Long

    ' Create a tiling object in order to create the background:
    Set cT = New cTile
    With cT
        .Initialise Me
        .FileName = App.Path & "\bck_001.bmp"
    End With
    
    ' Create sprites:
    CreateSpriteResource cResShip, App.Path & "\m_sh.bmp", 12, 1, &HFF00&
    CreateSprite cResShip, cShip
        
    CreateSpriteResource cResRockxL, App.Path & "\a_exlr.bmp", 1, 1, &HFF00&
    ReDim cRocks(1 To (4 + 8 + 16 + 32)) As cSprite
    For i = 1 To 4
        CreateSprite cResRockxL, cRocks(i)
        InitRockPosition cRocks(i)
    Next i
    CreateSpriteResource cResRockL, App.Path & "\a_large.bmp", 1, 1, &HFF00&
    For i = 1 To 8
        CreateSprite cResRockL, cRocks(i + 4)
        InitRockPosition cRocks(i + 4)
    Next i
    CreateSpriteResource cResRockM, App.Path & "\a_med.bmp", 1, 1, &HFF00&
    For i = 1 To 16
        CreateSprite cResRockM, cRocks(i + 12)
        InitRockPosition cRocks(i + 12)
    Next i
    CreateSpriteResource cResRockS, App.Path & "\a_small.bmp", 1, 1, &HFF00&
    For i = 1 To 32
        CreateSprite cResRockS, cRocks(i + 28)
        InitRockPosition cRocks(i + 28)
    Next i
            
    ' Create a bitmap on which to create the screen display
    ' offscreen.  This will be blitted from onto the screen
    ' to minimise flicker
    Set cStage = New cBitmap
    lW = Screen.Width \ Screen.TwipsPerPixelX
    lH = Screen.Height \ Screen.TwipsPerPixelY
    cStage.CreateAtSize lW, lH
    
    ' We tile the background bitmap into the stage bitmap
    ' to get some sort of background for the process:
    cT.TileDC cStage.hDC, lW, lH

    Me.Show
    Me.Refresh
    GameLoop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (m_bInGame) Then
        GameLoop
    End If
End Sub

