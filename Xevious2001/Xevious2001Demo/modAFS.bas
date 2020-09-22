Attribute VB_Name = "modAFS"
' This code was written by Chris Yates.  I am freely distributing this code under one condition.
' If you use this code in your program please give me credit and e-mail me (cyates@neo.rr.com) and
' tell me about your program.  Thanks, and enjoy.

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&     ' This setting is not in your API viewer, not sure why.
                                        ' If you use SC_MOVE then the mouse moves to the title bar
                                        ' and then moves the form, which makes forms with no title bar
                                        ' to not work.
Public Const WM_SYSCOMMAND = &H112

Dim CurRgn, TempRgn As Long  ' Region variables

Public Function AutoFormShape(bg As Form, transColor)
Dim X, Y As Integer

CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)  ' Create base region which is the current whole window

While Y <= bg.ScaleHeight  ' Go through each column of pixels on form
    While X <= bg.ScaleWidth  ' Go through each line of pixels on form
        If GetPixel(bg.hdc, X, Y) = transColor Then  ' If the pixels color is the transparency color (bright purple is a good one to use)
            TempRgn = CreateRectRgn(X, Y, X + 1, Y + 1)  ' Create a temporary pixel region for this pixel
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
            DeleteObject (TempRgn)  ' Delete the temporary pixel region and clear up very important resources
        End If
        X = X + 1
    Wend
        Y = Y + 1
        X = 0
Wend
success = SetWindowRgn(bg.hwnd, CurRgn, True)  ' Finally set the windows region to the final product
DeleteObject (CurRgn)  ' Delete the now un-needed base region and free resources

End Function

' This code is (C)1999 Chris Yates

