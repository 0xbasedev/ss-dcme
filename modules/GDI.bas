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

' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

Public Sub LoadPic(Pic As PictureBox, path As String)
    Dim Token As Long
    ' Initialise GDI+
    Token = InitGDIPlus
    Pic = LoadPictureGDIPlus(path)
    FreeGDIPlus Token
End Sub
' Initialises GDI Plus
Public Function InitGDIPlus() As Long
    Dim Token As Long
    Dim gdipInit As GdiplusStartupInput

    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional width As Long = -1, Optional height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim hDC As Long
    Dim hBitmap As Long
    Dim Img As Long

    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If

    ' Calculate picture's width and height if not specified
    If width = -1 Or height = -1 Then
        GdipGetImageWidth Img, width
        GdipGetImageHeight Img, height
    End If

    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, width, height

    ' Resize the picture
    gdipResize Img, hDC, width, height, RetainRatio
    GdipDisposeImage Img

    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
End Function

' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, width As Long, height As Long)
    Dim hBrush As Long

    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(width, height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, width, height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
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

    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic

    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight

        OrRatio = OrWidth / OrHeight
        DesRatio = width / height

        ' Calculate destination coordinates
        destWidth = IIf(DesRatio < OrRatio, width, height * OrRatio)
        destHeight = IIf(DesRatio < OrRatio, width / OrRatio, height)
        DestX = (width - destWidth) / 2
        DestY = (height - destHeight) / 2

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, destWidth, destHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, width, height
    End If
    GdipDeleteGraphics Graphics
End Sub

' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As GUID
    Dim Pic As PICTDESC
    Dim IPic As IPicture

    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46

    ' Fill Pic with necessary parts
    Pic.size = Len(Pic)        ' Length of structure
    Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    Pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function

' Returns a resized version of the picture
Public Function Resize(Handle As Long, PicType As PictureTypeConstants, width As Long, height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim Img As Long
    Dim hDC As Long
    Dim hBitmap As Long
    Dim WmfHeader As wmfPlaceableFileHeader

    ' Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
        GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
        FillInWmfHeader WmfHeader, width, height
        GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
        GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
        ' Does not return a valid Image object
        GdipCreateBitmapFromHICON Handle, Img
    End Select

    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, BackColor, width, height
        gdipResize Img, hDC, width, height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        Set Resize = CreatePicture(hBitmap)
    End If
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, width As Long, height As Long)
    WmfHeader.BoundingBox.Right = width
    WmfHeader.BoundingBox.Bottom = height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
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
    bgPen = CreateSolidBrush(color) ' CreatePen(0, 1, color)

    Dim fillarea As RECT
    With fillarea
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    
    Dim ret As Long
    
    ret = FillRect(hDC, fillarea, bgPen)

'            frmGeneral.Label6.Caption = "fill: " & ret
'        If ret = 0 Then MsgBox "fillrect failed!" & ret
    
    
    
    DeleteObject bgPen

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

Sub DrawImagePreview(ByRef srcPic As PictureBox, ByRef srcshape As shape, ByRef destpic As PictureBox, ByRef destshape As shape, ByVal BackColor As Long)
    
    
    Call DrawImagePreviewCoords(srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, destpic.hDC, destshape.Left, destshape.Top, destshape.width, destshape.height, BackColor)
    
    destpic.Refresh

End Sub

