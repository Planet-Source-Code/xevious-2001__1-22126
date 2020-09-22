VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   360
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dx As New DirectX7
Dim dd As DirectDraw7


'these surfaces hold the window
Dim DDS_primary As DirectDrawSurface7
Dim ddsd_primary As DDSURFACEDESC2
Dim DDS_back As DirectDrawSurface7
Dim ddsd_back As DDSURFACEDESC2

Dim DDS_Sfondo As DirectDrawSurface7
Dim ddsd_Sfondo As DDSURFACEDESC2


'a clipper for windowed mode
Dim ddClipper As DirectDrawClipper

'these surfaces hold our sprites
Dim DDS_run As DirectDrawSurface7
Dim ddsd_run As DDSURFACEDESC2
Dim DDS_other As DirectDrawSurface7
Dim ddsd_other As DDSURFACEDESC2
Dim r As RECT 'handy to have!

'hold the current screen mode
Dim smode As Integer

'hold the quit flag
Dim bquit



Private Sub Command1_Click()
    New_Game (0)
End Sub

Private Sub Form_Load()

    r.Top = 0
    r.Left = 0
    r.Right = 640
    r.Bottom = 480

End Sub
Public Sub New_Game(screenmode As Integer)
    'Form1.Hide

    'resize the form to the right size
    'Form2.Move 0, 0, 640 * Screen.TwipsPerPixelX, 480 * Screen.TwipsPerPixelY

    'remember the screenmode
    smode = screenmode
    
    'create the DirectDraw object
    Set dd = dx.DirectDrawCreate("")
    
    If screenmode = 0 Then
        '***************************************
        'windowed
        '***************************************
        'make your application a happy normal window
        Call dd.SetCooperativeLevel(Picture1.hWnd, DDSCL_NORMAL)
        
        'Indicate that the ddsCaps member is valid in this type
        ddsd_primary.lFlags = DDSD_CAPS
        'This surface is the primary surface (what is visible to the user)
        ddsd_primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        'Your creating the primary surface now with the surface description you just set
        Set DDS_primary = dd.CreateSurface(ddsd_primary)
        
        'allocate the clipper
        Set ddClipper = dd.CreateClipper(0)
        ddClipper.SetHWnd Picture1.hWnd
        DDS_primary.SetClipper ddClipper
        
        'This is going to be a plain off-screen surface
        ddsd_back.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        'tell create we want to set the width and height & caps
        ddsd_back.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        'at 640 by 480 in size
        ddsd_back.lHeight = 480
        ddsd_back.lWidth = 640
        'Now we create the 640x480 off-screen surface
        Set DDS_back = dd.CreateSurface(ddsd_back)
        'colour it in white
        DDS_back.BltColorFill r, RGB(255, 255, 255)
        Dim imgDim As ImgDimType
        
        a = getImgDim(App.Path & "\Sprites.bmp", imgDim, "")
        ddsd_Sfondo.lFlags = DDSD_CAPS
        ddsd_Sfondo.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        
        'tell create we want to set the width and height & caps
        'at 640 by 480 in size
        ddsd_Sfondo.lHeight = 480 'imgDim.height
        ddsd_Sfondo.lWidth = 640 'imgDim.width
        'Now we create the 640x480 off-screen surface
        Set DDS_Sfondo = dd.CreateSurfaceFromFile(App.Path & "\Sprites.bmp", ddsd_Sfondo)
        
        
        
    Else '***************************************
        ' full-screen!!
        '***************************************
        'make your application full-screen
        Call dd.SetCooperativeLevel(Picture1.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
        
        'set the screen mode
        dd.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT

        'get the screen surface and create a back buffer too
        ddsd_primary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        ddsd_primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
        ddsd_primary.lBackBufferCount = 1
        Set DDS_primary = dd.CreateSurface(ddsd_primary)
        
        'now grab the back surface (from the flipping chain)
        Dim caps As DDSCAPS2
        caps.lCaps = DDSCAPS_BACKBUFFER
        Set DDS_back = DDS_primary.GetAttachedSurface(caps)
                
    End If
    
    'show this form
    'Form2.Show
    
    
    '***************************************
    'load surfaces and any data
    'loaddata
    
    '***************************************
    'begin the main rendering loop
    renderloop
End Sub

Sub renderloop()
'the rectangles for windowed mode
Dim r2 As RECT

Do
    'draw current game screen
    drawframe

    'flip the double buffered surfaces
    If smode = 0 Then 'windowed
        dx.GetWindowRect Picture1.hWnd, r2
        r2.Top = r2.Top '+ 22
        'DDS_back.Blt
        'DDS_back.Blt r, DDS_Sfondo, r, DDBLTFAST_NOCOLORKEY + DDBLTFAST_WAIT
        DDS_primary.Blt r, DDS_Sfondo, r, DDBLTFAST_NOCOLORKEY + DDBLTFAST_WAIT
    Else 'full-screen
        DDS_primary.Flip Nothing, DDFLIP_WAIT
    End If
    
    'make time for other things
    DoEvents
    'Picture1.Refresh
Loop Until bquit
End Sub


Sub drawframe()
    Dim r2 As RECT
    r2.Top = 0
    r2.Left = 0
    r2.Right = 150
    r2.Bottom = 150
    'draw the current frame of action
    DDS_back.BltFast 10, 10, DDS_Sfondo, r2, DDBLTFAST_SRCCOLORKEY
End Sub


Private Sub Form_Unload(Cancel As Integer)
    bquit = True
End Sub

Private Function GetRect(ByVal n As Integer) As RECT
    GetRect.Top = Int(n / 10) * 32
    GetRect.Left = Int(n Mod 10) * 32
    GetRect.Bottom = GetRect.Top + 32
    GetRect.Right = GetRect.Left + 32
End Function

