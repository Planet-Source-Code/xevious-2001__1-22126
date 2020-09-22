VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Move(ByVal Index As Integer)
End Sub
Public Sub Collision(ByVal Index As Integer, ByVal CollidedIndex As Integer)
End Sub
Public Sub Show()
End Sub
Public Sub Animation(ByVal Index As Integer)
End Sub
Public Sub PrepareToStart()
End Sub
