Attribute VB_Name = "SpriteBas"
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
        
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

'This was put together by Andrew Heinlein (Mouse)
'mouse@theblackhand.net
'I Found the below function in C++ while surfing on the web
'and changed it over to Visual Basic API cause i thought
'it would be nice to have around... Transparent Bitmaps... with simple API
'who woulda thought....


'This will return TRUE if it succeeds. If you are using Win2000/NT you
'can even use GetLastError to give a better error than just FALSE.
'99.9% of the time, if it returns false, its because a parameter is wrong

'BOOL TransparentBlt(
'  HDC hdcDest,        // handle to destination DC
'  int nXOriginDest,   // x-coord of destination upper-left corner
'  int nYOriginDest,   // y-coord of destination upper-left corner
'  int nWidthDest,     // width of destination rectangle
'  int hHeightDest,    // height of destination rectangle
'  HDC hdcSrc,         // handle to source DC
'  int nXOriginSrc,    // x-coord of source upper-left corner
'  int nYOriginSrc,    // y-coord of source upper-left corner
'  int nWidthSrc,      // width of source rectangle
'  int nHeightSrc,     // height of source rectangle
'  UINT crTransparent  // color to make transparent
');

Public Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, ByVal hdcSrc As Long, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Long) As Boolean


Declare Function timeGetTime Lib "winmm.dll" () As Long


Function Mask(PicSrc As PictureBox, picDEST As PictureBox, bColor As OLE_COLOR)
    Dim looper As Long
    Dim looper2 As Long
    Dim bColor2 As OLE_COLOR
    picDEST.Cls
    For looper = 0 To PicSrc.ScaleHeight
    picDEST.Refresh
        For looper2 = 0 To PicSrc.ScaleWidth
            If PicSrc.Point(looper2, looper) = bColor Then
                bColor2 = RGB(255, 255, 255)
            Else
                bColor2 = RGB(0, 0, 0)
            End If
            SetPixel picDEST.hdc, looper2, looper, bColor2
        Next looper2
    Next looper
    picDEST.Refresh
End Function
Function Sprite(PicSrc As PictureBox, picDEST As PictureBox, bColor As OLE_COLOR)
    Dim looper As Long
    Dim looper2 As Long
    Dim bColor2 As OLE_COLOR
    picDEST.Cls
    For looper = 0 To PicSrc.ScaleHeight
    picDEST.Refresh
        For looper2 = 0 To PicSrc.ScaleWidth
            If PicSrc.Point(looper2, looper) = bColor Then
                bColor2 = RGB(0, 0, 0)
            Else
                bColor2 = GetPixel(PicSrc.hdc, looper2, looper)
            End If
            SetPixel picDEST.hdc, looper2, looper, bColor2
        Next looper2
    Next looper
    picDEST.Refresh
End Function
