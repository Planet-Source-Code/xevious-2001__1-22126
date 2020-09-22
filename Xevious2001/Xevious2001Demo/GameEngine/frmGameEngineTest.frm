VERSION 5.00
Begin VB.Form frmGameEngineTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin Project1.ctlGameEngine GameEngine1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _extentx        =   5106
      _extenty        =   4683
   End
End
Attribute VB_Name = "frmGameEngineTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    GameEngine1.LoadGraphics App.Path + "\..\backgroundbn.bmp", App.Path + "\..\sprites.bmp"
End Sub

Private Sub Timer1_Timer()
    GameEngine1.TimeClick
End Sub
