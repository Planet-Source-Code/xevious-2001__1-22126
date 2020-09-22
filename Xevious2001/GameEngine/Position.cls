VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim AbsX As Single
Dim AbsY As Single
Dim RelX As Single
Dim RelY As Single

Public Enum dnzDirection
    eUp = 1
    eRight = 2
    eDown = 4
    eUpDown = 5
    eLeft = 8
    eLeftRight = 10
    eCoveringFlag = 16     'Used to verify covering
End Enum

Public Sub setRelativePosition(ByVal X As Single, ByVal Y As Single, ByVal Region As RectRegion)
    RelX = X
    RelY = Y
    AbsX = Region.Left + X
    AbsY = Region.Top + Y
End Sub
Public Sub setAbsolutePosition(ByVal X As Single, ByVal Y As Single, ByVal Region As RectRegion)
    AbsX = X
    AbsY = Y
    RelX = X - Region.Left
    RelY = Y - Region.Top
End Sub
Public Property Get AbsoluteX() As Single
    AbsoluteX = AbsX
End Property
Public Property Get AbsoluteY() As Single
    AbsoluteY = AbsY
End Property
Public Property Get RelativeX() As Single
    RelativeX = RelX
End Property
Public Property Get RelativeY() As Single
    RelativeY = RelY
End Property
Public Sub CopyPosition(ByVal SourcePos As Position, ByVal Region As RectRegion)
    setAbsolutePosition SourcePos.AbsoluteX, SourcePos.AbsoluteY, Region
End Sub
