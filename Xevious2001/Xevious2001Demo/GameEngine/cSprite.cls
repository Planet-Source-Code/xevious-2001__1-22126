VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Memory DC & Bitmap to hold the sprite
Private m_hdcSprite As Long
Private m_hBmpSprite As Long
Private m_hBmpSpriteOld As Long

' Memory DC & Bitmap to hold the mask element of the sprite
Private m_hdcMask As Long
Private m_hBmpMask As Long
Private m_hBmpMaskOld As Long

' Memory DC to build up the sprite and draw it:
Private m_hDCBuild As Long
Private m_hBmpBuild As Long
Private m_hBmpBuildOld As Long

' Variables to store sprite frames, current cell and cell size:
Private m_lDX As Long
Private m_lDY As Long
Private m_lCX As Long
Private m_lCY As Long


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
    ByVal hDC As Long _
    ) As Long
' Creates a bitmap in memory:
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long _
    ) As Long
' Places a GDI object into DC, returning the previous one:
Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long _
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
    ByVal hDC As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
    ) As Long
' Create a brush of a certain colour:
Private Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
    ) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Function GDIMakeDCAndBitmap( _
        ByVal bMono As Boolean, _
        ByRef hDC As Long, _
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
    hDC = 0: hBMP = 0: hBmpOld = 0
    ' Create the DC from the basis DC:
    If (bMono) Then
        lCDC = 0
    Else
        lhWnd = GetDesktopWindow()
        lCDC = GetDC(lhWnd)
    End If
    hDC = CreateCompatibleDC(lCDC)
    If (bMono) Then
        lCDC = hDC
    End If
    
    If (hDC <> 0) Then
        ' If we get one, then time to make the bitmap:
        hBMP = CreateCompatibleBitmap(lCDC, lDX, lDY)
        ' If we succeed in creating the bitmap:
        If (hBMP <> 0) Then
            ' Select the bitmap into the memory DC and
            ' store the bitmap that was there before (need
            ' to do this because you need to select this
            ' bitmap back into the DC before deleting
            ' the new Bitmap):
            hBmpOld = SelectObject(hDC, hBMP)
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
        ByRef hDC As Long, _
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
    hDC = 0: hBMP = 0: hBmpOld = 0
    
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
        hDC = CreateCompatibleDC(hDCBasis)
        If (hDC <> 0) Then
            ' If DC Is created, select the bitmap into it:
            hBmpOld = SelectObject(hDC, hBMP)
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
        ByRef hDC As Long, _
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
    hDC = 0: hBMP = 0: hBmpOld = 0
        
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
    
        hDC = CreateCompatibleDC(hDCBasis)
        If (hDC <> 0) Then
            ' If we get one, then time to make the bitmap:
            Dim tBM As Bitmap
            GetObjectAPI oPic.Handle, Len(tBM), tBM
            
            hBMP = CreateCompatibleBitmap(hDCBasis, tBM.bmWidth, tBM.bmHeight)
            If (hBMP <> 0) Then
                hBmpOld = SelectObject(hDC, hBMP)
                
                BitBlt hDC, 0, 0, tBM.bmWidth, tBM.bmHeight, hDCTemp, 0, 0, SRCCOPY
                
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
        ByRef hDC As Long, _
        ByRef hBMP As Long, _
        ByVal hBmpOld As Long _
    )
' **********************************************************
' GDI Helper function: Goes through the steps required
' to clear up a bitmap within a DC.
' **********************************************************
    ' If we have a valid DC:
    If (hDC <> 0) Then
        ' If there is a valid bitmap in it:
        If (hBMP <> 0) Then
            ' Select the original bitmap into the DC:
            SelectObject hDC, hBmpOld
            ' Now delete the unreferenced bitmap:
            DeleteObject hBMP
            ' Byref so set the value to invalid BMP:
            hBMP = 0
        End If
        ' Delete the memory DC:
        DeleteObject hDC
        ' Byref so set the value to invalid DC:
        hDC = 0
    End If
End Sub




Public Sub TransparentDraw( _
        ByVal hDCDest As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Boolean = True _
    )
Dim lDX As Long
Dim lDY As Long
Dim lDC As Long
Dim loX As Long
Dim loY As Long

    ' Get the position of the sprite cell within the data:
    pGetXYForCell lCell, lDX, lDY
    
    ' Determine where to draw to:
    If (bUseBuildDC) Then
        ' Now we take a copy of the background we are drawing on
        ' into the work DC:
        lDC = m_hDCBuild
        BitBlt lDC, 0, 0, m_lDX, m_lDY, hDCDest, lX, lY, SRCCOPY
        loX = 0
        loY = 0
    Else
        lDC = hDCDest
        loX = lX
        loY = lY
    End If
    
    ' We OR the mask into place, i.e. wherever the mask is black
    ' there is no effem_cTile, but wherever it is white the background
    ' goes completely white:
    BitBlt lDC, loX, loY, m_lDX, m_lDY, m_hdcMask, lDX, lDY, SRCPAINT
    ' Now we AND the sprite into position: wherever the sprite is white,
    ' there is no effem_cTile, but over the whitened mask area the sprite
    ' shows as normal:
    BitBlt lDC, loX, loY, m_lDX, m_lDY, m_hdcSprite, lDX, lDY, SRCAND
    
    If (bUseBuildDC) Then
        ' Now we can blit the combination of the background with the
        ' transparently drawn area onto the Destination:
        BitBlt hDCDest, lX, lY, m_lDX, m_lDY, m_hDCBuild, 0, 0, SRCCOPY
    End If
    
End Sub

Public Function CreateFromPicture( _
        ByRef oPic As StdPicture, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByRef oMaskPic As StdPicture = Nothing, _
        Optional ByVal lTransColor As Long = &HFFFFFF _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean
    
    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' First create the sprite bitmap:
    If (GDILoadPictureIntoDC( _
            False, _
            oPic, _
            m_hdcSprite, _
            m_hBmpSprite, _
            m_hBmpSpriteOld _
            )) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpSprite, Len(tBM), tBM
        
        If (oMaskPic Is Nothing) Then
            ' Create a mask based on white areas of the
            ' original sprite:
            bContinue = pbCreateMaskFromSprite( _
                tBM.bmWidth, _
                tBM.bmHeight, _
                lTransColor _
                )
        Else
            ' We load the specified mask image:
            bContinue = GDILoadPictureIntoDC( _
                True, _
                oMaskPic, _
                m_hdcMask, _
                m_hBmpMask, _
                m_hBmpMaskOld)
        End If
        
        If (bContinue) Then
            ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            m_lDX = tBM.bmWidth \ m_lCX
            m_lDY = tBM.bmHeight \ m_lCY
            ' Make a working DC to draw the sprite in:
            If (GDIMakeDCAndBitmap( _
                False, _
                m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld, _
                m_lDX, m_lDY)) Then
                CreateFromPicture = True
            Else
                Debug.Print "Failed to create working DC/Bitmap"
            End If
        Else
            Debug.Print "Failed to create mask DC/Bitmap"
        End If
    Else
        Debug.Print "Failed to create sprite DC/Bitmap"
    End If
    
End Function

Public Function CreateFromFile( _
        ByVal sSpriteBitmap As String, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByVal sMaskBitmap As String = "", _
        Optional ByVal lTransColor As Long = &HFFFFFF _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean

    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' First create the sprite bitmap:
    If (GDILoadBitmapIntoDC( _
            False, _
            sSpriteBitmap, _
            m_hdcSprite, _
            m_hBmpSprite, _
            m_hBmpSpriteOld _
            )) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpSprite, Len(tBM), tBM
        
        If (sMaskBitmap = "") Then
            ' Create a mask based on white areas of the
            ' original sprite:
            bContinue = pbCreateMaskFromSprite( _
                tBM.bmWidth, _
                tBM.bmHeight, _
                lTransColor _
                )
        Else
            ' We load the specified mask image:
            bContinue = GDILoadBitmapIntoDC( _
                True, _
                sMaskBitmap, _
                m_hdcMask, _
                m_hBmpMask, _
                m_hBmpMaskOld)
        End If
        
        If (bContinue) Then
            ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            m_lDX = tBM.bmWidth \ m_lCX
            m_lDY = tBM.bmHeight \ m_lCY
            ' Make a working DC to draw the sprite in:
            If (GDIMakeDCAndBitmap( _
                False, _
                m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld, _
                m_lDX, m_lDY)) Then
                CreateFromFile = True
            Else
                Debug.Print "Failed to create working DC/Bitmap"
            End If
        Else
            Debug.Print "Failed to create mask DC/Bitmap"
        End If
    Else
        Debug.Print "Failed to create sprite DC/Bitmap"
    End If
    
End Function
Private Function pbCreateMaskFromSprite( _
        ByVal bmWidth As Long, _
        ByVal bmHeight As Long, _
        ByVal lTransColor As Long _
    ) As Boolean
Dim lOrigColor As Long
Dim bTransColorIsWhite As Boolean
    
    ' Now create a mask Object.  Here we want
    ' to just copy the bitmap into a monochrome
    ' DC, then invert it:
    If (GDIMakeDCAndBitmap( _
        True, _
        m_hdcMask, m_hBmpMask, m_hBmpMaskOld, _
        bmWidth, bmHeight)) Then
        
        If (lTransColor = &HFFFFFF) Then
            bTransColorIsWhite = True
        End If
        
        If (bTransColorIsWhite) Then
            ' Copy the Bitmap from the source, inverting it
            ' as we go:
            BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, &HFF0062 ' Whiteness
            BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, SRCINVERT
            pbCreateMaskFromSprite = True
        Else
            ' We need to make an non-inverse mask first.  This
            ' is ORed with the original sprite to make the area
            ' with transparent colour = white, then it the mask is
            ' inverted.
            Dim hDCTemp As Long, hBmpTemp As Long, hBmpTempOld As Long
            
            If (GDIMakeDCAndBitmap( _
                True, _
                hDCTemp, hBmpTemp, hBmpTempOld, _
                bmWidth, bmHeight)) Then
                
                ' Set the back colour for the sprite to the
                ' transparent colour (this means it is set to
                ' black when we bitblt to the mono DC):
                lOrigColor = SetBkColor(m_hdcSprite, lTransColor)
                ' Make non-inverse mask:
                BitBlt hDCTemp, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, SRCCOPY
                ' Return back colour of sprite
                SetBkColor m_hdcSprite, lOrigColor
                ' Make sprite white in lTransColor region:
                BitBlt m_hdcSprite, 0, 0, bmWidth, bmHeight, hDCTemp, 0, 0, SRCPAINT
                ' Invert the mask:
                BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, &HFF0062 ' Whiteness
                BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, hDCTemp, 0, 0, SRCINVERT
                
                ' Clear up temporary bitmap:
                GDIClearDCBitmap hDCTemp, hBmpTemp, hBmpTempOld
            
                pbCreateMaskFromSprite = True
            End If
        End If
    End If

End Function
Property Get CellWidth() As Long
    CellWidth = m_lDX
End Property
Property Get CellHeight() As Long
    CellHeight = m_lDY
End Property
Public Sub DirectBltSprite( _
        ByVal lHDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long _
    )
Dim lDX As Long
Dim lDY As Long
    pGetXYForCell lCell, lDX, lDY
    BitBlt lHDC, lX, lY, m_lDX, m_lDY, m_hdcSprite, lDX, lDY, SRCCOPY
End Sub
Private Sub pGetXYForCell( _
        ByVal lCell As Long, _
        ByRef lX As Long, _
        ByRef lY As Long _
    )
    'lX = ((lCell - 1) Mod m_lCX) * m_lDX
    lX = ((lCell) Mod m_lCX) * m_lDX
    'lY = ((lCell - 1) \ m_lCX) * m_lDY
    lY = ((lCell) \ m_lCX) * m_lDY
End Sub

Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    GDIClearDCBitmap m_hdcSprite, m_hBmpSprite, m_hBmpSpriteOld
    GDIClearDCBitmap m_hdcMask, m_hBmpMask, m_hBmpMaskOld
    GDIClearDCBitmap m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld
End Sub

Private Sub Class_Initialize()
    Debug.Print "m_cSpriteBitmaps:Initialise"
End Sub

Private Sub Class_Terminate()
    Debug.Print "m_cSpriteBitmaps:Terminate"
    ClearUp
End Sub



