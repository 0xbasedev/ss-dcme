Attribute VB_Name = "GDI_Pack"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC
    size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type PWMFRect16
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key As Long
    hMf As Integer
    BoundingBox As PWMFRect16
    Inch As Integer
    Reserved As Long
    CheckSum As Integer
End Type


' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long


' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal filename As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)


Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, _
                                        ByVal X As Long, ByVal Y As Long, _
                                        ByVal lpString As String, _
                                        ByVal nCount As Long) As Long
                                        
' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

' Logical Font
Private Const TEXT_TRANSPARENT = 1
Private Const TEXT_OPAQUE = 2

Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const FONT_SIZE = 12
Private Const NO_ERROR = 0
Private Const ANSI_CHARSET = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0
Private Const LOGPIXELSY = 90
Private Const TRANSPARENT = 1

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type




Public Sub LoadPic(Pic As PictureBox, path As String)
          Dim Token As Long
          ' Initialise GDI+
10        Token = InitGDIPlus
20        Pic = LoadPictureGDIPlus(path)
30        FreeGDIPlus Token
End Sub
' Initialises GDI Plus
Public Function InitGDIPlus() As Long
          Dim Token As Long
          Dim gdipInit As GdiplusStartupInput

10        gdipInit.GdiplusVersion = 1
20        GdiplusStartup Token, gdipInit, ByVal 0&
30        InitGDIPlus = Token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
10        GdiplusShutdown Token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional width As Long = -1, Optional height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
          Dim hDC As Long
          Dim hBitmap As Long
          Dim Img As Long

          ' Load the image
10        If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
20            Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
30            Exit Function
40        End If

          ' Calculate picture's width and height if not specified
50        If width = -1 Or height = -1 Then
60            GdipGetImageWidth Img, width
70            GdipGetImageHeight Img, height
80        End If

          ' Initialise the hDC
90        InitDC hDC, hBitmap, BackColor, width, height

          ' Resize the picture
100       gdipResize Img, hDC, width, height, RetainRatio
110       GdipDisposeImage Img

          ' Get the bitmap back
120       GetBitmap hDC, hBitmap

          ' Create the picture
130       Set LoadPictureGDIPlus = CreatePicture(hBitmap)
End Function

' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, width As Long, height As Long)
          Dim hBrush As Long

          ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
10        hDC = CreateCompatibleDC(ByVal 0&)
20        hBitmap = CreateBitmap(width, height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
30        hBitmap = SelectObject(hDC, hBitmap)
40        hBrush = CreateSolidBrush(BackColor)
50        hBrush = SelectObject(hDC, hBrush)
60        PatBlt hDC, 0, 0, width, height, PATCOPY
70        DeleteObject SelectObject(hDC, hBrush)
End Sub

' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hDC As Long, width As Long, height As Long, Optional RetainRatio As Boolean = False)
          Dim Graphics As Long    ' Graphics Object Pointer
          Dim OrWidth As Long   ' Original Image Width
          Dim OrHeight As Long    ' Original Image Height
          Dim OrRatio As Double    ' Original Image Ratio
          Dim DesRatio As Double  ' Destination rect Ratio
          Dim DestX As Long    ' Destination image X
          Dim DestY As Long    ' Destination image Y
          Dim destWidth As Long     ' Destination image Width
          Dim destHeight As Long      ' Destination image Height

10        GdipCreateFromHDC hDC, Graphics
20        GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic

30        If RetainRatio Then
40            GdipGetImageWidth Img, OrWidth
50            GdipGetImageHeight Img, OrHeight

60            OrRatio = OrWidth / OrHeight
70            DesRatio = width / height

              ' Calculate destination coordinates
80            destWidth = IIf(DesRatio < OrRatio, width, height * OrRatio)
90            destHeight = IIf(DesRatio < OrRatio, width / OrRatio, height)
100           DestX = (width - destWidth) / 2
110           DestY = (height - destHeight) / 2

120           GdipDrawImageRectRectI Graphics, Img, DestX, DestY, destWidth, destHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
130       Else
140           GdipDrawImageRectI Graphics, Img, 0, 0, width, height
150       End If
160       GdipDeleteGraphics Graphics
End Sub

' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
10        hBitmap = SelectObject(hDC, hBitmap)
20        DeleteDC hDC
End Sub

' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
          Dim IID_IDispatch As GUID
          Dim Pic As PICTDESC
          Dim IPic As IPicture

          ' Fill in OLE IDispatch Interface ID
10        IID_IDispatch.Data1 = &H20400
20        IID_IDispatch.Data4(0) = &HC0
30        IID_IDispatch.Data4(7) = &H46

          ' Fill Pic with necessary parts
40        Pic.size = Len(Pic)        ' Length of structure
50        Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
60        Pic.hBmp = hBitmap         ' Handle to bitmap

          ' Create the picture
70        OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
80        Set CreatePicture = IPic
End Function

' Returns a resized version of the picture
Public Function Resize(Handle As Long, PicType As PictureTypeConstants, width As Long, height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
          Dim Img As Long
          Dim hDC As Long
          Dim hBitmap As Long
          Dim WmfHeader As wmfPlaceableFileHeader

          ' Determine pictyre type
10        Select Case PicType
          Case vbPicTypeBitmap
20            GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
30        Case vbPicTypeMetafile
40            FillInWmfHeader WmfHeader, width, height
50            GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
60        Case vbPicTypeEMetafile
70            GdipCreateMetafileFromEmf Handle, False, Img
80        Case vbPicTypeIcon
              ' Does not return a valid Image object
90            GdipCreateBitmapFromHICON Handle, Img
100       End Select

          ' Continue with resizing only if we have a valid image object
110       If Img Then
120           InitDC hDC, hBitmap, BackColor, width, height
130           gdipResize Img, hDC, width, height, RetainRatio
140           GdipDisposeImage Img
150           GetBitmap hDC, hBitmap
160           Set Resize = CreatePicture(hBitmap)
170       End If
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, width As Long, height As Long)
10        WmfHeader.BoundingBox.Right = width
20        WmfHeader.BoundingBox.Bottom = height
30        WmfHeader.Inch = 1440
40        WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub

























Sub DrawLine(hDC As Long, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, color As Long)
    Dim bgPen As Long
    Dim orgPt As POINTAPI
    
    bgPen = CreatePen(0, 1, color)
    
    Dim oldobj As Long
    oldobj = SelectObject(hDC, bgPen)
    
    MoveToEx hDC, x1, y1, orgPt
    LineTo hDC, x2, y2
    MoveToEx hDC, orgPt.X, orgPt.Y, orgPt
    
    
    SelectObject hDC, oldobj
    DeleteObject bgPen
    
End Sub

Sub DrawFilledRectangle(hDC As Long, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, color As Long)
          Dim bgPen As Long
10        bgPen = CreateSolidBrush(color) ' CreatePen(0, 1, color)

          Dim fillarea As RECT
20        With fillarea
30            .Left = Left
40            .Top = Top
50            .Right = Right
60            .Bottom = Bottom
70        End With
          
          Dim ret As Long
          
80        ret = FillRect(hDC, fillarea, bgPen)

'            frmGeneral.Label6.Caption = "fill: " & ret
'        If ret = 0 Then MsgBox "fillrect failed!" & ret
          
          
          
90        DeleteObject bgPen

End Sub

Sub DrawRectangle(hDC As Long, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, color As Long)
    Dim bgPen As Long
    Dim orgPt As POINTAPI
    
    bgPen = CreatePen(0, 1, color)
    
    Dim oldobj As Long
    oldobj = SelectObject(hDC, bgPen)
    
    MoveToEx hDC, Left, Top, orgPt
    LineTo hDC, Left, Bottom
    LineTo hDC, Right, Bottom
    LineTo hDC, Right, Top
    LineTo hDC, Left, Top
    
    
    'Reset stuff
    MoveToEx hDC, orgPt.X, orgPt.Y, orgPt
    SelectObject hDC, oldobj
    DeleteObject bgPen


End Sub

Sub ChangeTextColor(hDC As Long, color As Long)
    Call SetTextColor(hDC, color)
    Call SetBkMode(hDC, TEXT_TRANSPARENT)
End Sub


Sub ChangeTextSize(hDC As Long, size As Integer)
    Dim hFont As Long, hOldFont As Long
    Dim lf As LOGFONT
    
'    With lf
    lf.lfHeight = size
'    End With
    
    hFont = CreateFontIndirectA(lf)
    
    hOldFont = SelectObject(hDC, hFont)
    
    
    DeleteObject hOldFont
End Sub

Sub PrintText(hDC As Long, Text As String, X As Long, Y As Long)
    TextOut hDC, X, Y, Text, Len(Text)
End Sub


Sub DrawImagePreviewCoords(srcDC As Long, SrcX As Long, SrcY As Long, srcWidth As Long, srcHeight As Long, destDC As Long, DestX As Long, DestY As Long, destWidth As Long, destHeight As Long, BackColor As Long)


    Call DrawFilledRectangle(destDC, CInt(DestX), CInt(DestY), CInt(DestX + destWidth), CInt(DestY + destHeight), BackColor)
    
    
    If srcWidth > srcHeight Then
        'Resize considering width

        If srcWidth = destWidth Then
            'Same size, use bitblt
            BitBlt destDC, DestX, DestY + (destHeight \ 2) - (srcHeight \ 2), srcWidth, srcHeight, srcDC, SrcX, SrcY, vbSrcCopy
            
        ElseIf srcWidth < destWidth Then
            'Source is smaller, use pixel resize
            SetStretchBltMode destDC, COLORONCOLOR
            StretchBlt destDC, DestX, DestY + (destHeight \ 2) - ((srcHeight / (srcWidth / destWidth)) \ 2), destWidth, srcHeight / (srcWidth / destWidth), srcDC, SrcX, SrcY, srcWidth, srcHeight, vbSrcCopy
        
        Else
            'Source is larger, use halftone resize
            SetStretchBltMode destDC, HALFTONE
            StretchBlt destDC, DestX, DestY + (destHeight \ 2) - ((srcHeight / (srcWidth / destWidth)) \ 2), destWidth, srcHeight / (srcWidth / destWidth), srcDC, SrcX, SrcY, srcWidth, srcHeight, vbSrcCopy
        End If
    Else
        If srcHeight = destHeight Then
            'Same size, use bitblt
            BitBlt destDC, DestX + (destWidth \ 2) - (srcWidth \ 2), DestY, srcWidth, srcHeight, srcDC, SrcX, SrcY, vbSrcCopy
            
        ElseIf srcHeight < destHeight Then
            'Source is smaller, use pixel resize
            SetStretchBltMode destDC, COLORONCOLOR
            StretchBlt destDC, DestX + (destWidth \ 2) - ((srcWidth / (srcHeight / destHeight)) \ 2), DestY, srcWidth / (srcHeight / destHeight), destHeight, srcDC, SrcX, SrcY, srcWidth, srcHeight, vbSrcCopy
        
        Else
            'Source is larger, use halftone resize
            SetStretchBltMode destDC, HALFTONE
            StretchBlt destDC, DestX + (destWidth \ 2) - ((srcWidth / (srcHeight / destHeight)) \ 2), DestY, srcWidth / (srcHeight / destHeight), destHeight, srcDC, SrcX, SrcY, srcWidth, srcHeight, vbSrcCopy
        End If
    End If
End Sub

Sub DrawImagePreview(ByRef srcPic As PictureBox, ByRef srcshape As Shape, ByRef destPic As PictureBox, ByRef destshape As Shape, ByVal BackColor As Long)
    
    
    Call DrawImagePreviewCoords(srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, destPic.hDC, destshape.Left, destshape.Top, destshape.width, destshape.height, BackColor)
    
    destPic.Refresh

End Sub




'Sub ExportImageList(ByRef il As ImageList, filename As String)
'    Dim curX As Integer
'    curX = 0
'
'    Dim output As New clsDisplayLayer
'    Dim listimg As ListImage
'
'    output.BackColor = il.MaskColor
'
'    Call output.Resize(il.ListImages.count * il.imageWidth, il.imageHeight, False)
'    Call output.Cls
'
'    For Each listimg In il.ListImages
'
'        Call listimg.Draw(output.hDC, curX, 0)
'
'        curX = curX + (il.imageWidth * Screen.TwipsPerPixelX)
'
'    Next
'
'    Call output.SaveToFile(filename, False)
'End Sub
