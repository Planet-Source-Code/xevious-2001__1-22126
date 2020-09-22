VERSION 5.00
Begin VB.UserControl GameEngine 
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   ScaleHeight     =   3645
   ScaleWidth      =   5310
   Begin BackgroundCtl.Sprite Sprite2 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin BackgroundCtl.Sprite Sprite1 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin BackgroundCtl.Background Background1 
      Height          =   3015
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   3836
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "GameEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Time As Timer           'Used to get timer
Public Picture As PictureBox    'The picture used to represent game
Public Background As PictureBox    'The picture used to represent game
Public collEnemy As Collection  'Collects all enemies


Public Sub StartGame()
    Timer1.Enabled = True
End Sub
Public Sub EndGame()
    Timer1.Enabled = False
End Sub
Public Sub LoadGraphics()
    Set Sprite1.Picture = Background1.Picture                  'Specify sprite back ground
    Sprite1.LoadSprite (App.Path + "\explosion32.gif") 'Load Sprite
    Sprite1.DimX = 32                               'Define image dimension
    Sprite1.DimY = 32

    Set Sprite2.Picture = Background1.Picture                  'Specify sprite back ground
    Sprite2.LoadSprite (App.Path + "\SpritesXevious.BMP") 'Load Sprite
    Sprite2.DimX = 32                               'Define image dimension
    Sprite2.DimY = 32

    Background1.LoadScrollingImage (App.Path + "\background.gif")

End Sub


Public Sub Timer()
    'This is used to control each tic
    
    
End Sub

Private Sub UserControl_Initialize()
    Set Picture = Picture1
    Set Background = Background1.Picture
    'Set Timer = Timer1
    

End Sub
