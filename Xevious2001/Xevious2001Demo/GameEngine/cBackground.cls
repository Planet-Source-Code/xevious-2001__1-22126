VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Memory DC & Bitmap to hold the bitmap
Private m_hdcBitmap As Long
Private m_hBmpBitmap As Long
Private m_hBmpBitmapOld As Long
' Width & Height of the bitmap
Private m_lDX As Long
Private m_lDY As Long

' API Declares:

' This is most useful but Win32 only.  Particularly try the
' LOADMAP3DCOLORS for a quick way to sort out those
' embarassing gray backgrounds in your fixed bitmaps!
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal un1 As Long, _
    ByVal n1 As Long, ByVal n2 As Long, _
    ByVal un2 As Long _
    ) As Long
    
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const LR_COLOR = &H2
Private Const LR_COPYDELETEORG = &H8
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_COPYRETURNORG = &H4
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_MONOCHROME = &H1
Private Const LR_SHARED = &H8000

' Creates a memory DC
Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long _
    ) As Long
' Creates a bitmap in memory:
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long _
    ) As Long
' Places a GDI object into DC, returning the previous one:
Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long _
    ) As Long
' Deletes a GDI object:
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long _
    ) As Long
' Copies Bitmaps from one DC to another, can also perform
' raster operations during the transfer:
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long
Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal nWidthDest As Long, _
    ByVal nHeightDest As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal dwRop As Long _
    ) As Long
Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086
Private Const SRCINVERT = &H660046

' Structure used to hold bitmap information about Bitmaps
' created using GDI in memory:
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
' Get information relating to a GDI Object
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
    ) As Long
' The traditional rectangle structure:
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Fills a rectangle in a DC with a specified brush
Private Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
    ) As Long
' Create a brush of a certain colour:
Private Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
    ) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Function GDIMakeDCAndBitmap( _
        ByVal bMono As Boolean, _
        ByRef hdc As Long, _
        ByRef hBMP As Long, _
        ByRef hBmpOld As Long, _
        ByVal lDX As Long, _
        ByVal lDY As Long _
    ) As Boolean
' **********************************************************
' GDI Helper function: Makes a bitmap of a specified size
' and creates a DC to hold it.
' **********************************************************
Dim lCDC As Long
Dim lhWnd As Long

    ' Initialise byref variables:
    hdc = 0: hBMP = 0: hBmpOld = 0
    ' Create the DC from the basis DC:
    If (bMono) Then
        lCDC = 0
    Else
        lhWnd = GetDesktopWindow()
        lCDC = GetDC(lhWnd)
    End If
    hdc = CreateCompatibleDC(lCDC)
    If (bMono) Then
        lCDC = hdc
    End If
    
    If (hdc <> 0) Then
        ' If we get one, then time to make the bitmap:
        hBMP = CreateCompatibleBitmap(lCDC, lDX, lDY)
        ' If we succeed in creating the bitmap:
        If (hBMP <> 0) Then
            ' Select the bitmap into the memory DC and
            ' store the bitmap that was there before (need
            ' to do this because you need to select this
            ' bitmap back into the DC before deleting
            ' the new Bitmap):
            hBmpOld = SelectObject(hdc, hBMP)
            ' Success:
            GDIMakeDCAndBitmap = True
        End If
    End If
    
    If Not (bMono) Then
        ReleaseDC lhWnd, lCDC
    End If

End Function
Private Function GDILoadBitmapIntoDC( _
        ByVal bMono As Boolean, _
        ByVal sFileName As String, _
        ByRef hdc As Long, _
        ByRef hBMP As Long, _
        ByRef hBmpOld As Long _
    ) As Boolean
' **********************************************************
' GDI Helper function: Loads a bitmap from file and selects
' it into a memory DC.
' **********************************************************
Dim hInst As Long
Dim hDCBasis As Long
Dim lhWnd As Long

    ' Initialise byref variables:
    hdc = 0: hBMP = 0: hBmpOld = 0
    
    ' Now load the sprite bitmap:
    hInst = App.hInstance
    ' This is the quick, direct way where we don't get
    ' any extra copies of the bitmaps, as compared to
    ' using the VB picture object:
    hBMP = LoadImage(hInst, sFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If (hBMP <> 0) Then
        ' Create a DC to hold the sprite, and select
        ' the sprite into it:
        If (bMono) Then
            hDCBasis = 0
        Else
            lhWnd = GetDesktopWindow()
            hDCBasis = GetDC(lhWnd)
        End If
        hdc = CreateCompatibleDC(hDCBasis)
        If (hdc <> 0) Then
            ' If DC Is created, select the bitmap into it:
            hBmpOld = SelectObject(hdc, hBMP)
            GDILoadBitmapIntoDC = True
        End If
        If Not (bMono) Then
            ReleaseDC lhWnd, hDCBasis
        End If
    End If

End Function
Private Function GDILoadPictureIntoDC( _
        ByVal bMono As Boolean, _
        ByRef oPic As StdPicture, _
        ByRef hdc As Long, _
        ByRef hBMP As Long, _
        ByRef hBmpOld As Long _
    ) As Boolean
' **********************************************************
' GDI Helper function: Creates a memory DC containing a new
' copy of bitmap from a StdPicture.
' **********************************************************
Dim hInst As Long
Dim hDCBasis As Long
Dim lhWnd As Long
Dim hDCTemp As Long
Dim hBmpTemp As Long
Dim hBmpTempOld As Long

    ' Initialise byref variables:
    hdc = 0: hBMP = 0: hBmpOld = 0
        
    ' Create a DC to hold the sprite, and select
    ' the sprite into it:
    If (bMono) Then
        hDCBasis = 0
    Else
        lhWnd = GetDesktopWindow()
        hDCBasis = GetDC(lhWnd)
    End If
    hDCTemp = CreateCompatibleDC(hDCBasis)
    If (bMono) Then
        hDCBasis = hDCTemp
    End If
    
    If (hDCTemp <> 0) Then
        hBmpTempOld = SelectObject(hDCTemp, oPic.Handle)
    
        hdc = CreateCompatibleDC(hDCBasis)
        If (hdc <> 0) Then
            ' If we get one, then time to make the bitmap:
            Dim tBM As Bitmap
            GetObjectAPI oPic.Handle, Len(tBM), tBM
            
            hBMP = CreateCompatibleBitmap(hDCBasis, tBM.bmWidth, tBM.bmHeight)
            If (hBMP <> 0) Then
                hBmpOld = SelectObject(hdc, hBMP)
                
                BitBlt hdc, 0, 0, tBM.bmWidth, tBM.bmHeight, hDCTemp, 0, 0, SRCCOPY
                
                GDILoadPictureIntoDC = True
            End If
        End If
        
        SelectObject hDCTemp, hBmpTempOld
        DeleteObject hDCTemp
        
    End If
    If Not (bMono) Then
        ReleaseDC lhWnd, hDCBasis
    End If

End Function

Private Sub GDIClearDCBitmap( _
        ByRef hdc As Long, _
        ByRef hBMP As Long, _
        ByVal hBmpOld As Long _
    )
' **********************************************************
' GDI Helper function: Goes through the steps required
' to clear up a bitmap within a DC.
' **********************************************************
    ' If we have a valid DC:
    If (hdc <> 0) Then
        ' If there is a valid bitmap in it:
        If (hBMP <> 0) Then
            ' Select the original bitmap into the DC:
            SelectObject hdc, hBmpOld
            ' Now delete the unreferenced bitmap:
            DeleteObject hBMP
            ' Byref so set the value to invalid BMP:
            hBMP = 0
        End If
        ' Delete the memory DC:
        DeleteObject hdc
        ' Byref so set the value to invalid DC:
        hdc = 0
    End If
End Sub


Public Function CreateAtSize( _
        ByVal lDX As Long, _
        ByVal lDY As Long _
    ) As Boolean

    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' Create a DC at size to hold the bitmap:
    If (GDIMakeDCAndBitmap( _
        False, _
        m_hdcBitmap, _
        m_hBmpBitmap, _
        m_hBmpBitmapOld, _
        lDX, lDY)) Then
        m_lDX = lDX
        m_lDY = lDY
        CreateAtSize = True
    End If
    
End Function
Public Function CreateFromFile( _
        ByVal sBitmapBitmap As String _
    ) As Boolean
Dim tBM As Bitmap

    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' Load the Bitmap bitmap:
    If (GDILoadBitmapIntoDC( _
        False, _
        sBitmapBitmap, _
        m_hdcBitmap, _
        m_hBmpBitmap, _
        m_hBmpBitmapOld)) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpBitmap, Len(tBM), tBM
        ' Store dimensions:
        m_lDX = tBM.bmWidth
        m_lDY = tBM.bmHeight
        CreateFromFile = True
    End If
    
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim tBM As Bitmap
    ClearUp
    If (GDILoadPictureIntoDC(False, picThis, m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld)) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpBitmap, Len(tBM), tBM
        ' Store dimensions:
        m_lDX = tBM.bmWidth
        m_lDY = tBM.bmHeight
        CreateFromPicture = True
    End If
End Function
Property Get Width() As Long
    Width = m_lDX
End Property
Property Get Height() As Long
    Height = m_lDY
End Property
Property Get hdc() As Long
    hdc = m_hdcBitmap
End Property
Public Sub RenderBitmap( _
        ByVal hdcDest As Long, _
        ByVal xDest As Long, _
        ByVal yDest As Long, _
        Optional ByVal OpCode As RasterOpConstants = vbSrcCopy _
    )
    BitBlt _
        hdcDest, _
        xDest, yDest, m_lDX, m_lDY, _
        m_hdcBitmap, _
        0, 0, _
        OpCode
End Sub
Public Sub RenderBitmapWindowInPicture( _
        ByVal hdcDest As Long, _
        ByVal xDest As Long, _
        ByVal yDest As Long, _
        ByVal nWidthDest As Long, _
        ByVal nHeightDest As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        Optional ByVal OpCode As RasterOpConstants = vbSrcCopy _
    )
    StretchBlt _
        hdcDest, _
        xDest, yDest, nWidthDest, nHeightDest, _
        m_hdcBitmap, _
        xSrc, ySrc, nWidthDest, nHeightDest, _
        OpCode
End Sub
Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    GDIClearDCBitmap m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld
End Sub
Private Sub Class_Terminate()
    ClearUp
End Sub




