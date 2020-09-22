VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Top As Long
Public Left As Long
Public Width As Long
Public Height As Long

Public Sub MoveUp(ByVal lDeltaY As Long)
    Top = Top - lDeltaY
End Sub
Public Sub MoveLeft(ByVal lDeltaX As Long)
    Left = Left - lDeltaX
End Sub
Public Sub SetRect(Optional ByVal lLeft As Long = 0, Optional ByVal lTop As Long = 0, Optional ByVal lWidth As Long = 0, Optional ByVal lHeight As Long = 0)
    Left = lLeft
    Top = lTop
    Width = lWidth
    Height = lHeight
End Sub
