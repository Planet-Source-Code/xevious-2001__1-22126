VERSION 5.00
Begin VB.UserControl ctlGameEngine 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   ScaleHeight     =   5550
   ScaleWidth      =   7920
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   120
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.Label lblScoreLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblScrollText 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GameEngine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1320
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   2145
      End
   End
End
Attribute VB_Name = "ctlGameEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Time As Timer           'Used to get timer
'Public sp As New ISprite
Dim rDimensionRatio As Single            'Used to reduce bitmap dimension
Public ScrollingRegion As New RectRegion
Public Picture As PictureBox    'The picture used to represent game
'Dim BackGroundCtl As Background    'The picture used to represent game
Public collEnemy As New Collection  'Collects all enemies
Public collDeletedSprites As New Collection
Public UserAirMissile As Integer        'Count number of air missiles shoot
Public UserGroundMissile As Integer     'Count number of ground missiles shoot
Public ScrollDownSpeed As Integer
Public ScrollLeftSpeed As Integer
Public ScoreLabel As Label
'Public ScollTextLabel As Label
Public UserSpriteName As String
Public UserSprite As Object
Public BackgroundImageDimension As New RectRegion
Dim BackgroundText As String, TimeBackgroundText As Integer

Public MousePos As New Position
Dim GameScore As Long
Public Energy As Long
Public Difficulty As Long

Dim Hiscore(10) As String
Dim ScrollHiscore As Boolean

Dim MapLines() As String
Dim MapIndex As Long        'The last map line read

Dim Background As New cBackground
Dim Sprite As New cSprite
Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event EndScrollingRegion(Direction As dnzDirection)
Public Event UserCollision()
Public Event NewMapObject(ObjectMapParam As String)





Public Sub LoadGraphics(ByVal sBackgroundImagePath As String, ByVal sSpriteImagePath As String)
    Sprite.CreateFromFile sSpriteImagePath, 10, 17, , 65280
    
    'Set Sprite1.Picture = Background1.Picture                  'Specify sprite back ground
    'Sprite1.LoadSprite (sSpriteImagePath) 'Load Sprite
    'Sprite1.DimX = 32                               'Define image dimension
    'Sprite1.DimY = 32

    If sBackgroundImagePath <> "" Then
        'Background1.LoadScrollingImage (sBackgroundImagePath)
        Background.CreateFromFile sBackgroundImagePath
    End If
    BackgroundImageDimension.SetRect 0, 0, Background.Width, Background.Height
End Sub

Public Sub TimeClick()
    'Background1.ScrollDown (ScrollDownSpeed)
    
    ScrollDown (ScrollDownSpeed)
    AnalyzeMap
    'Ordina in base a ZOrder
    Dim obj As Object
    OrderSpriteByZOrder

    'Dim obj As Object
    i = 1
    For Each obj In collEnemy
        obj.Move (i)
        obj.Show
        i = i + 1
    Next obj
    CollisionDetection
    DeleteSprites
    ShowFrame
    ShowLabel BackgroundText, TimeBackgroundText

End Sub
Public Sub ScrollDown(ByVal nPix As Long)
    'Background1.ScrollDown (ScrollDownSpeed)
    With ScrollingRegion
        .MoveUp -nPix * rDimensionRatio
        'If sPathFileLoaded <> "" Then
        'PasteScrollingRegion
        Background.RenderBitmapWindowInPicture Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, .Left, .Top
        If .Top + nPix < 0 And .Top > nPix * 2 Then
            RaiseEvent EndScrollingRegion(eUp)
        End If
        If .Top + nPix + .Height > Picture1.Height Then
            RaiseEvent EndScrollingRegion(eDown)
        End If
    
            
        'Else
            'Picture2.Cls
        'End If
    End With

End Sub


Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
    Set Picture = Picture1
    rDimensionRatio = 1
    Picture1.Left = 0
    Picture1.Top = 0
    'Background1.DimensionRatio = 1
    'Sprite1.DimensionRatio = 1
    'Set BackGroundCtl = Background1
    'Set ScrollingRegion = Background1.ScrollingRegion
    Set ScoreLabel = lblScoreLabel
    Set ScollTextLabel = lblScrollText
    UserControl_Resize
    'Set BackgroundImageDimension = Background1.BackgroundPictureDimension
    
End Sub

Private Sub UserControl_Resize()
    'Picture1 is always large as control!
    Picture1.Width = Width
    Picture1.Height = Height
    ScoreLabel.Width = Picture1.Width
    lblScrollText.Width = Picture1.ScaleWidth
    lblScrollText.Height = Picture1.ScaleHeight
    ScrollingRegion.Height = Picture1.ScaleHeight
    ScrollingRegion.Width = Picture1.ScaleWidth
    'Background Same dimension of picture1 (The same of designed control!)
    'Background1.Width = Width * rDimensionRatio
    'Background1.Height = Height * rDimensionRatio
    
End Sub
'Public Property Let DimensionRatio(ByVal sDimensionRatio As Single)
'    rDimensionRatio = sDimensionRatio
'    UserControl_Resize
'End Property
'Public Property Get DimensionRatio() As Single
'    DimensionRatio = rDimensionRatio
'End Property
'Public Property Let BackgroundDimensionRatio(ByVal sDimensionRatio As Single)
    'rBackgroundDimensionRatio = sDimensionRatio
'    Background1.DimensionRatio = sDimensionRatio
'End Property
'Public Property Let SpriteDimensionRatio(ByVal sDimensionRatio As Single)
    'rBackgroundDimensionRatio = sDimensionRatio
'    Sprite1.DimensionRatio = sDimensionRatio
'End Property
Public Sub ShowFrame()
    'Picture1.PaintPicture Background1.Picture.Image, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Background1.Picture.ScaleWidth, Background1.Picture.ScaleHeight
    Picture1.Refresh
End Sub
Public Sub CreateNewSprite(objSprite As Object, Optional ByVal sName As String = "")
    On Error Resume Next
    If sName <> "" Then
        collEnemy.Add objSprite, sName
        If Err > 0 Then
            collEnemy.Remove sName
            collEnemy.Add objSprite, sName
        End If
    Else
        collEnemy.Add objSprite ', sName
    End If
End Sub
Public Sub RemoveSprite(ByVal Index As Integer)
    collEnemy.Remove (Index)
    
End Sub

Public Sub PasteSprite(ByVal RelX As Long, ByVal RelY As Long, Optional ByVal SpriteNum As Integer = 0)
    'SpriteNumX = SpriteNum Mod Sprite1.NumSpriteX
    'SpriteNumY = Int(SpriteNum / Sprite1.NumSpriteX)
    'Sprite1.PasteSprite RelX, RelY, SpriteNumX, SpriteNumY     'Specify where to paste and the index of image to paste
    Sprite.TransparentDraw Picture1.hdc, RelX, RelY, SpriteNum, True
End Sub
    
Public Property Get Score() As Long
    Score = GameScore
End Property
Public Property Let Score(ByVal lScore As Long)
    GameScore = lScore
End Property

Sub CollisionDetection()
    Dim a As dnzDirection
    If collEnemy.Count > 1 Then
        For i = 1 To collEnemy.Count - 1
            With collEnemy(i).SpriteRegion
                On Error Resume Next
                If collEnemy(i).ZOrder > 9 Then
                    For o = i + 1 To collEnemy.Count
                        If collEnemy(i).SpriteRegion.Width * collEnemy(o).SpriteRegion.Width <> 0 Then
                            a = RectCovering(collEnemy(i).SpritePosition, collEnemy(i).SpriteRegion, collEnemy(o).SpritePosition, collEnemy(o).SpriteRegion)
                            If (a And eCoveringFlag) <> 0 Then
                                collEnemy(i).Collision i, o
                                collEnemy(o).Collision o, i
                            Else
                                a = RectCovering(collEnemy(o).SpritePosition, collEnemy(o).SpriteRegion, collEnemy(i).SpritePosition, collEnemy(i).SpriteRegion)
                                If (a And eCoveringFlag) <> 0 Then
                                    collEnemy(i).Collision i, o
                                    collEnemy(o).Collision o, i
                                End If
                            End If
                        End If
                    Next o
                End If
                On Error GoTo 0
            End With
        Next i
    End If
End Sub
Public Function RectCovering(CoveredRegionPosition As Position, CoveredRegion As RectRegion, CoveringRegionPosition As Position, CoveringRegion As RectRegion) As dnzDirection
    Dim DirectionCovering As dnzDirection
    If ((CoveringRegionPosition.AbsoluteX + CoveringRegion.Left) < (CoveredRegionPosition.AbsoluteX + CoveredRegion.Left)) And ((CoveringRegionPosition.AbsoluteX + CoveringRegion.Left + CoveringRegion.Width) > (CoveredRegionPosition.AbsoluteX + CoveredRegion.Left)) Then
        DirectionCovering = DirectionCovering Or eLeft
    End If
    If ((CoveringRegionPosition.AbsoluteY + CoveringRegion.Top) < (CoveredRegionPosition.AbsoluteY + CoveredRegion.Top)) And (CoveringRegionPosition.AbsoluteY + CoveringRegion.Top + CoveringRegion.Height) > (CoveredRegionPosition.AbsoluteY + CoveredRegion.Top) Then
        DirectionCovering = DirectionCovering Or eUp
    End If
    If (CoveringRegionPosition.AbsoluteY + CoveringRegion.Top) < (CoveredRegionPosition.AbsoluteY + CoveredRegion.Top + CoveredRegion.Height) And (CoveringRegionPosition.AbsoluteY + CoveringRegion.Top + CoveringRegion.Height) > (CoveredRegionPosition.AbsoluteY + CoveredRegion.Top + CoveredRegion.Height) Then
        DirectionCovering = DirectionCovering Or eDown
    End If
    If (CoveringRegionPosition.AbsoluteX + CoveringRegion.Left) < (CoveredRegionPosition.AbsoluteX + CoveredRegion.Left + CoveredRegion.Width) And (CoveringRegionPosition.AbsoluteX + CoveringRegion.Left + CoveringRegion.Width) > (CoveredRegionPosition.AbsoluteX + CoveredRegion.Left + CoveredRegion.Width) Then
        DirectionCovering = DirectionCovering Or eRight
    End If
    If (DirectionCovering And eLeftRight) <> 0 And (DirectionCovering And eUpDown) <> 0 Then
        DirectionCovering = DirectionCovering Or eCoveringFlag
    End If
    RectCovering = DirectionCovering
End Function
Private Sub DeleteSprites()
    EndFor = collDeletedSprites.Count
    If EndFor = 0 Then Exit Sub
    Dim NotOrdered As Boolean
    With collDeletedSprites
        NotOrdered = True
        Do While NotOrdered
            NotOrdered = False
        'For i = 1 To EndFor - 1
            For o = 1 To EndFor - 1
            'For o = i + 1 To EndFor
                If collDeletedSprites(o) > collDeletedSprites(o + 1) Then
                    tmp = collDeletedSprites.Item(o)
                    .Remove (o)
                    .Add tmp, , , collDeletedSprites.Count
                    'o = o - 1
                    NotOrdered = True
                End If
            Next o
        Loop
        'Next i
        OldDel = -1
        For i = EndFor To 1 Step -1
            NewDel = .Item(i)
            If NewDel <> OldDel Then
                collEnemy.Remove NewDel
                OldDel = NewDel
            End If
            .Remove i
        Next i
    End With
End Sub
Public Function IsSpriteInScrollingRegion(Sprite As Object) As Boolean
    IsSpriteInScrollingRegion = True
    With Sprite
        If .SpritePosition.RelativeX + 32 < 0 Then
            IsSpriteInScrollingRegion = False
            Exit Function
        End If
        If .SpritePosition.RelativeY + 32 < 0 Then
            IsSpriteInScrollingRegion = False
            Exit Function
        End If
        If .SpritePosition.RelativeX > ScrollingRegion.Width Then
            IsSpriteInScrollingRegion = False
            Exit Function
        End If
        If .SpritePosition.RelativeY > ScrollingRegion.Height Then
            IsSpriteInScrollingRegion = False
            Exit Function
        End If
    End With
End Function
Public Sub SpritePathMove(ByVal Time As Single, ByVal PathType As Integer, IncX As Single, IncY As Single)
    Dim Speed As Single
    Speed = 1
    T = Time
    Pi = 3.14
    rad = T * (Pi / 90)
    r = Int(ScrollingRegion.Width / 5)
    Select Case PathType
        Case 1  'Horizontal
            IncX = Speed
        Case 2  'Vertical
            IncY = Speed
        Case 3  'Circular
            OldRad = (T - 1) * (Pi / 90)
            IncX = r * (Cos(rad) - Cos(OldRad))
            IncY = r * (Sin(rad) - Sin(OldRad))
        Case 4  'Square
            T = T Mod 100
            If T < 25 Then
                IncX = Speed
                IncY = 0
            ElseIf T < 50 Then
                IncX = 0
                IncY = Speed
            ElseIf T < 75 Then
                IncX = -Speed
                IncY = 0
            Else
                IncX = 0
                IncY = -Speed
            End If
        Case 5  'Square 45°
            T = T Mod (4 * r)
            If T < r Then
                IncX = Speed
                IncY = Speed
            ElseIf T < 2 * r Then
                IncX = -Speed
                IncY = Speed
            ElseIf T < 3 * r Then
                IncX = -Speed
                IncY = -Speed
            Else
                IncX = Speed
                IncY = -Speed
            End If
        Case Else
            IncX = Speed
            IncY = Speed
    End Select
End Sub
Public Sub UserSpriteCollision()
    RaiseEvent UserCollision
End Sub
Public Function UserSpritePosition() As Position
    Dim Pippo As New Position
    Static x As Single
    Static y As Single
    Dim OldX As Single
    Dim OldY As Single
    OldX = x
    OldY = y
    On Error GoTo NoUserSprite
    x = collEnemy(UserSpriteName).SpritePosition.RelativeX
    y = collEnemy(UserSpriteName).SpritePosition.RelativeY
AllOk:
    Pippo.setRelativePosition x, y, ScrollingRegion
    Set UserSpritePosition = Pippo
    Exit Function
NoUserSprite:
    x = OldX
    y = OldY
    Pippo.setRelativePosition x, y, ScrollingRegion
    Set UserSpritePosition = Pippo
    Resume AllOk
End Function
Public Sub UserFire(ByVal FireType As Integer)
    On Error Resume Next
    If collEnemy(UserSpriteName).SpriteName = UserSpriteName Then
    End If
    If Err = 0 Then
        UserSprite.Fire (FireType)
    End If
End Sub
Public Sub ClearAllSprite()
    For i = 1 To collEnemy.Count
        collDeletedSprites.Add i
    Next i
End Sub
Private Sub OrderSpriteByZOrder()
    Dim obj As Object
    For i = 1 To collEnemy.Count
        For o = 1 To collEnemy.Count - 1
            If collEnemy.Item(o).ZOrder > collEnemy.Item(o + 1).ZOrder Then
                Set obj = collEnemy.Item(o)
                ObjKey$ = ""
                If obj.SpriteName = "Solvalou" Then
                    ObjKey$ = "Solvalou"
                End If
                collEnemy.Remove (o)
                If ObjKey$ <> "" Then
                    collEnemy.Add obj, ObjKey$, , o
                Else
                    collEnemy.Add obj, , , o
                End If
            End If
        Next o
    Next i
End Sub
Public Sub LoadMap(ByVal sPathFile As String)
    ff = FreeFile
    i = 0
    Open sPathFile For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, a$
            a$ = Trim(a$)
            If a$ <> "" And Mid$(a$, 1, 1) <> "'" Then
                'Is not empty and is not a comment
                ReDim Preserve MapLines(i)
                MapLines(i) = a$
                i = i + 1
            End If
        Loop
    Close #ff
    MapIndex = 0
End Sub
Private Sub AnalyzeMap()
    Dim MapParam() As Variant
    Dim OtherObj As Boolean
    Dim FirstNotLoadedObject As String
    On Error Resume Next
    If UBound(MapLines) > 0 Then
    End If
    If Err = 0 Then
        OtherObj = True
        Do While OtherObj And MapIndex <= UBound(MapLines)
            FirstNotLoadedObject = MapLines(MapIndex)
            OtherObj = False
            If (Val(Split(FirstNotLoadedObject, ",")(0)) > ScrollingRegion.Left And Val(Split(FirstNotLoadedObject, ",")(0)) < ScrollingRegion.Left + ScrollingRegion.Width) Or LCase(Split(FirstNotLoadedObject, ",")(0)) = "x" Then
                If (Val(Split(FirstNotLoadedObject, ",")(1)) > ScrollingRegion.Top And Val(Split(FirstNotLoadedObject, ",")(1)) < ScrollingRegion.Top + ScrollingRegion.Height) Or LCase(Split(FirstNotLoadedObject, ",")(1)) = "x" Then
                    RaiseEvent NewMapObject(FirstNotLoadedObject)
                    MapIndex = MapIndex + 1
                    OtherObj = True
                End If
            End If
        Loop
    End If
    On Error GoTo 0

End Sub

Public Sub SetStartingPosition(ByVal lLeft As Long, ByVal lTop As Long)
    'Background1.SetStartingPosition lLeft, lTop
    ScrollingRegion.Left = lLeft
    ScrollingRegion.Top = lTop
    
End Sub

Sub ShowLabel(ByVal sString As String, ByVal nDecSec As Integer)
    Static T As Long
    Static lString As String
    If T > nDecSec Then
        lblScrollText.Visible = False
        ScrollHiscore = False
    End If
    If Not ScrollHiscore Then
        If lString <> sString Then
            lblScrollText.Caption = sString
            lString = sString
            lblScrollText.Left = (Picture1.ScaleWidth - lblScrollText.Width) / 2
            lblScrollText.Top = (Picture1.ScaleHeight - lblScrollText.Height) / 2
            lblScrollText.Visible = True
            T = 0
        End If
    Else
        lString = sString
        lblScrollText = sString
        lblScrollText.Visible = True
        lblScrollText.Top = lblScrollText.Top - 2
        lblScrollText.Left = (Picture1.ScaleWidth - lblScrollText.Width) / 2
        If lblScrollText.Top - 2 < (Picture1.ScaleHeight - lblScrollText.Height) / 2 Then ScrollHiscore = False
        T = 0
    End If
    T = T + 1
End Sub
Public Sub SetBackGroundText(ByVal sString As String, ByVal nDecSec As Integer)
    BackgroundText = sString
    TimeBackgroundText = nDecSec
    ScrollHiscore = False

End Sub

Public Sub LoadHighScore(ByVal sPathFile As String)
    On Error Resume Next
    ff = FreeFile
    i = 0
    Open sPathFile For Input As #ff
        If Err > 0 Then Exit Sub
        Do While Not EOF(ff)
            Line Input #ff, a$
            Hiscore(i) = a$
            i = i + 1
        Loop
    Close #ff
End Sub
Public Sub UpdateHighScore(ByVal sUserName As String, ByVal nPoints As Long)
    If nPoints < Val(Hiscore(9)) Then Exit Sub
    Hiscore(10) = Format(nPoints, "000000") + "    " + sUserName
    'sort (Hiscore)
    
    Dim SortedArray As Boolean
    SortedArray = True
    
    Do
        SortedArray = True
        For loopcount = 0 To 9
            If Val(Hiscore(loopcount)) < Val(Hiscore(loopcount + 1)) Then
                SortedArray = False
                Call swap(Hiscore, loopcount, loopcount + 1)
            End If
        Next loopcount
    Loop Until SortedArray = True

End Sub
Private Sub sort(ByRef tmparray)
    Dim SortedArray As Boolean
    Dim Start, Finish As Integer
    SortedArray = True
    Start = LBound(tmparray)
    Finish = UBound(tmparray)


    Do
        SortedArray = True


        For loopcount = Start To Finish - 1


            If Val(tmparray(loopcount)) < Val(tmparray(loopcount + 1)) Then
                SortedArray = False
                Call swap(tmparray, loopcount, loopcount + 1)
            End If
        Next loopcount
    Loop Until SortedArray = True
End Sub
Private Sub swap(swparray, fpos, spos)
    Dim temp As Variant
    temp = swparray(fpos)
    swparray(fpos) = swparray(spos)
    swparray(spos) = temp
End Sub
Public Sub SaveHighScore(ByVal sFilePath As String)
    ff = FreeFile
    Open sFilePath For Output As #ff
        For i = 0 To 9
            Print #ff, Hiscore(i)
        Next i
    Close #ff
End Sub
Public Sub ShowHighScore()
    For i = 0 To 9
        a$ = a$ + Hiscore(i) + Chr$(10) + Chr$(13)
    Next i
    'lblScrollText = a$
    lblScrollText.Top = ScrollingRegion.Height
    SetBackGroundText a$, 200
    ScrollHiscore = True
End Sub
Public Function IsInHighScore(ByVal nPoints As Long)
    If nPoints > Val(Hiscore(9)) Then IsInHighScore = True
End Function
Public Function GetRandomNumber() As Double
    Static NewOrOld As Integer
    
    If NewOrOld = 0 Then
        a = timeGetTime() Mod 10000
        GetRandomNumber = a / 10001
        NewOrOld = NewOrOld + 1
    ElseIf NewOrOld = 1 Then
        GetRandomNumber = Rnd(timeGetTime)
        NewOrOld = NewOrOld + 1
    Else
        GetRandomNumber = Rnd(0)
        NewOrOld = 0
    End If
End Function



