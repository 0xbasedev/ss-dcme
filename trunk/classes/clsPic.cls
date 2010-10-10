VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 ' color table in RGBs

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long



Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private viewBMP As BITMAP

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SAFEARRAYBOUND
End Type
Private viewSA As SAFEARRAY1D


Private Data() As Byte
Dim Pic As PictureBox

Dim bitmapinfo24 As BITMAPINFO
Dim idc As Long
Dim ibitmap As Long

Dim r As RECT

Dim refreshIDC As Boolean
Dim bInit As Boolean

Sub Init(p As PictureBox, Optional keepContents As Boolean)
          'Initialise DMA Picture Access
          
10        Set Pic = p
          
          
20        With bitmapinfo24.bmiHeader
30            .biBitCount = 24
40            .biCompression = BI_RGB
50            .biPlanes = 1
60            .biSize = Len(bitmapinfo24.bmiHeader)
70            .biWidth = Pic.width ' viewBMP.bmWidth
80            .biHeight = Pic.height 'viewBMP.bmHeight
90        End With
          
          'create a dc and bitmap that we use to copy our DIB to, and be able to use it as a source for bitblt
          'clear old ones if already made
100       If idc <> 0 Then DeleteDC (idc)
110       If ibitmap <> 0 Then DeleteObject (ibitmap)
          
120       idc = CreateCompatibleDC(0) 'create a compatible dc
          'create a DIB section from the newly made idc
130       ibitmap = CreateDIBSection(idc, bitmapinfo24, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
140       SelectObject idc, ibitmap
          
150       If keepContents Then
              'bitblt hdc --> idc
160           BitBlt idc, 0, 0, Pic.width, Pic.height, Pic.hDC, 0, 0, vbSrcCopy
170       End If
          
          
180       Call CreatePicture(p)
          
          
190       GetObjectAPI p.Picture, Len(viewBMP), viewBMP   'retrieve bitmap information about p
          
200       With viewSA
210         .cbElements = 1
220         .cDims = 1
230         .Bounds(0).lLbound = 0
240         .Bounds(0).cElements = viewBMP.bmHeight * viewBMP.bmWidthBytes '+ viewBMP.bmWidth
            '.Bounds(1).lLbound = 0
            '.Bounds(1).cElements = viewBMP.bmWidthBytes
250         .pvData = viewBMP.bmBits
260       End With

          ' copy bitmap data into byte array
270       CopyMemory ByVal VarPtrArray(Data), VarPtr(viewSA), 4

280       r.Left = 0
290       r.Top = 0
300       r.Bottom = viewBMP.bmHeight
310       r.Right = viewBMP.bmWidth
          
          
320       If keepContents Then
              '    GetDIBits idc, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
              'append pixel array
330           GetDIBits idc, ibitmap, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
              
              'idc = pixel array
340           refreshIDC = False
350       End If
          
360       bInit = True
End Sub

Private Sub CreatePicture(p As PictureBox)
          'Creates a picture in memory
          
10        Set p.Picture = CreateDIBPicture(p.ScaleX(p.ScaleWidth, p.ScaleMode, vbPixels), p.ScaleY(p.ScaleHeight, p.ScaleMode, vbPixels), 24)
20        Exit Sub
          
      '    create a temp picture'
          Dim path As String
          Dim nr As Integer
        
30        nr = Rnd() * 10000
40        path = App.path & "\" & nr & ".bmp"
        
50        Do While Dir$(path) <> ""
60            nr = Rnd() * 10000
         '     path = App.path & "\" & nr & ".bmp"
70        Loop
80        Set p.Picture = p.Image
90        Call SavePicture(p.Picture, path)
100       p.Picture = LoadPicture(path)
        
110       Kill path
          
End Sub

Private Function CreateDIBPicture(ByVal pixelWidth As Long, ByVal pixelHeight As Long, ByVal BitsPerPixel As Long) As StdPicture
        'creates a DIB picture
        
        Dim IID_IDispatch As GUID
        Dim pbmp As PicBmp
        Dim Pic As StdPicture
          
10      With IID_IDispatch
20        .Data1 = &H20400
30        .Data4(0) = &HC0&
40        .Data4(7) = &H46&
50      End With
          
60      With pbmp
70        .size = Len(pbmp)
80        .Type = vbPicTypeBitmap
90        .hBmp = CreateDIB(pixelWidth, pixelHeight, BitsPerPixel)
100       .hPal = &H0&
110     End With
          
120     OleCreatePictureIndirect pbmp, IID_IDispatch, 1, Pic
130     Set CreateDIBPicture = Pic
        
End Function

Private Function CreateDIB(width As Long, height As Long, BitsPerPixel As Long) As Long
        'creates a DIB area
        
        Const BI_RGB As Long = 0
        Const DIB_RGB_COLORS As Long = 0

        Dim BMI As BITMAPINFO
        Dim stride As Long, ptrToBits As Long, hDC As Long

10      stride = (width * (BitsPerPixel \ 8) + 3) And &HFFFFFFFC
        
20      With BMI.bmiHeader
30        .biSize = 40
40        .biWidth = width
50        .biHeight = height
60        .biPlanes = 1
70        .biBitCount = BitsPerPixel
80        .biCompression = BI_RGB
90        .biSizeImage = stride * height
100     End With
        
110     hDC = CreateCompatibleDC(0)
120       CreateDIB = CreateDIBSection(hDC, BMI, DIB_RGB_COLORS, ptrToBits, 0, 0)
130     DeleteDC hDC
        
End Function


Sub stretchToDC(hDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, xSrc As Integer, ySrc As Integer, srcWidth As Integer, srcHeight As Integer, dwRop As Long)
          'Stretches DIB -> hdc
          
          'append dc
10        If refreshIDC Then SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
           
          'now bitblt idc --> hdc
20        StretchBlt hDC, X, Y, width, height, idc, xSrc, ySrc, srcWidth, srcHeight, dwRop
          
          'idc = pixel array
30        refreshIDC = False
End Sub

Sub stretchFromDC(hDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, xSrc As Integer, ySrc As Integer, srcWidth As Integer, srcHeight As Integer, dwRop As Long)
          'Stretches hdc --> DIB
          
          'make sure the idc is uptodate first
          'append dc
10        If refreshIDC Then SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
          
          'bitblt hdc --> idc
20        StretchBlt idc, X, Y, width, height, hDC, xSrc, ySrc, srcWidth, srcHeight, dwRop
          
          '    GetDIBits idc, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
          'append pixel array
30        GetDIBits idc, ibitmap, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
          
          'idc = pixel array
40        refreshIDC = False
End Sub

Sub setStretchMode(nStretchMode As Long)
          'sets the stretch mode of our hdc
          
10        Call SetStretchBltMode(idc, nStretchMode)
End Sub

Sub bltToDC(hDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, xSrc As Integer, ySrc As Integer, dwRop As Long)
          'blits DIB -> hdc
          
          'append dc
10        If refreshIDC Then SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
           
          'now bitblt idc --> hdc
20        BitBlt hDC, X, Y, width, height, idc, xSrc, ySrc, dwRop
          
          'idc = pixel array
30        refreshIDC = False
End Sub

Sub bltFromDC(hDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, xSrc As Integer, ySrc As Integer, dwRop As Long)
          'blits hdc -> DIB

          'make sure the idc is uptodate first
          'append dc
10        If refreshIDC Then SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
          
          'bitblt hdc --> idc
20        BitBlt idc, X, Y, width, height, hDC, xSrc, ySrc, dwRop
          
          '    GetDIBits idc, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
          'append pixel array
30        GetDIBits idc, ibitmap, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
          
          'idc = pixel array
40        refreshIDC = False
End Sub

Sub transToDC(hDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, xSrc As Integer, ySrc As Integer, color As Long)
          'blits DIB -> hdc
          
          'append dc
10        If refreshIDC Then SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
           
          'now bitblt idc --> hdc
20        TransparentBlt hDC, X, Y, width, height, idc, xSrc, ySrc, color
          
          'idc = pixel array
30        refreshIDC = False
End Sub

Sub transFromDC(hDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, xSrc As Integer, ySrc As Integer, color As Long)
          'blits hdc -> DIB

          'make sure the idc is uptodate first
          'append dc
10        If refreshIDC Then SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
          
          
          
          'bitblt hdc --> idc
20        TransparentBlt idc, X, Y, width, height, hDC, xSrc, ySrc, color
          
          '    GetDIBits idc, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
          'append pixel array
30        GetDIBits idc, ibitmap, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
          
          'idc = pixel array
40        refreshIDC = False
End Sub

Sub Refresh()
          'Refreshes the picture window
          
10        InvalidateRect Pic.hWnd, r, False
End Sub

Sub setPixel1D(idx As Long, r As Byte, g As Byte, b As Byte)
10        On Error Resume Next
20        Data(idx) = b
30        Data(idx + 1) = g
40        Data(idx + 2) = r
          
50        refreshIDC = True
End Sub



Sub SetPixel(X As Integer, Y As Integer, r As Byte, g As Byte, b As Byte)
              
          'Sets a pixel
10        On Error Resume Next
          
          Dim idx As Long
20        idx = ((viewBMP.bmHeight - 1) - Y) * viewBMP.bmWidthBytes + X * 3
          
30        Data(idx) = b
40        Data(idx + 1) = g
50        Data(idx + 2) = r
          
60        refreshIDC = True
End Sub

Sub setPixelLong(ByRef X As Integer, ByRef Y As Integer, ByRef color As Long)
10        Call SetPixel(X, Y, color And 255, (color And 65280) \ 256, (color And 16711680) \ 65535)
End Sub

Sub GetPixelRGB(ByRef X As Integer, ByRef Y As Integer, b() As Byte)
          Dim idx As Long
10        idx = ((viewBMP.bmHeight - 1) - Y) * viewBMP.bmWidthBytes + X * 3
              
20        CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(Data(idx)), 3

End Sub

Function GetPixel(ByRef X As Integer, ByRef Y As Integer) As Long
          Dim idx As Long
10        idx = ((viewBMP.bmHeight - 1) - Y) * viewBMP.bmWidthBytes + X * 3

20        GetPixel = RGB(Data(idx + 2), Data(idx + 1), Data(idx))
End Function

Sub Clear()
10        ZeroMemory ByVal VarPtr(Data(0)), UBound(Data) + 1
20        refreshIDC = True
          
End Sub

Sub Fill(grey As Byte)
10        FillMemory ByVal VarPtr(Data(0)), UBound(Data) + 1, ByVal grey
20        refreshIDC = True
End Sub

Sub fillarea(grey As Byte, Left As Long, Top As Long, Right As Long, Bottom As Long)
          Dim ptr As Long, row As Long, width As Long
          
10        ptr = VarPtr(Data(Top * viewBMP.bmWidthBytes + Left))
          
20        width = Right - Left + 1
          
30        For row = Top To Bottom
40            FillMemory ByVal ptr, width, grey
              
              
50            ptr = ptr + viewBMP.bmWidthBytes
60        Next row
          
70        refreshIDC = True
End Sub

Sub FillRGB(r As Byte, g As Byte, b As Byte)
          Dim i As Long, j As Long, idx As Long
          
10        For j = 0 To viewBMP.bmHeight - 1
20            idx = j * viewBMP.bmWidthBytes
              
30            For i = 0 To viewBMP.bmWidth - 1
40                Data(idx) = b
50                Data(idx + 1) = g
60                Data(idx + 2) = r

70                idx = idx + 3
80            Next
90        Next
          
100       refreshIDC = True
End Sub

Sub FillAreaRGB(r As Byte, g As Byte, b As Byte, Left As Long, Top As Long, Right As Long, Bottom As Long)
          Dim idx As Long, row As Long, col As Long, width As Long
          
10        idx = VarPtr(Data(Top * viewBMP.bmWidthBytes + Left))
          
20        width = Right - Left + 1
          
30        For row = Top To Bottom
40            idx = row * viewBMP.bmWidthBytes + Left
              
50            For col = Left To Right
                  
60                Data(idx) = b
70                Data(idx + 1) = g
80                Data(idx + 2) = r

90                idx = idx + 3
100           Next col
110       Next row
          
120       refreshIDC = True
End Sub


Private Sub UpdateIDC()
          'append dc
10        SetDIBitsToDevice idc, 0, 0, bitmapinfo24.bmiHeader.biWidth, bitmapinfo24.bmiHeader.biHeight, 0, 0, 0, bitmapinfo24.bmiHeader.biHeight, Data(0), bitmapinfo24, DIB_RGB_COLORS
20        refreshIDC = False
End Sub

Private Sub Class_Terminate()
10        Call Destroy
End Sub


'Private Sub TransparentBlt(dsthdc As Long, X As Integer, Y As Integer, width As Integer, height As Integer, sourcehDC As Long, xSrc As Integer, ySrc As Integer, transColor As Long)
'          Dim maskDC As Long      'DC for the mask
'          Dim tempDC As Long      'DC for temporary data
'          Dim hMaskBmpA As Long    'Bitmap for mask
'          Dim hTempBmpA As Long    'Bitmap for temporary data
'          Dim hMaskBmpB As Long    'Bitmap for mask
'          Dim hTempBmpB As Long    'Bitmap for temporary data
'          Dim tempbmp As Long
'          Dim tempbmp2 As Long
'          Dim srchDC As Long
'
'10        srchDC = CreateCompatibleDC(sourcehDC)
'20        tempbmp = CreateCompatibleBitmap(sourcehDC, width, height)
'30        tempbmp2 = SelectObject(srchDC, tempbmp)
'40        BitBlt srchDC, 0, 0, width, height, sourcehDC, xSrc, ySrc, vbSrcCopy
'
'          'First create some DC's. These are our gateways to assosiated bitmaps in RAM
'50        maskDC = CreateCompatibleDC(dsthdc)
'60        tempDC = CreateCompatibleDC(dsthdc)
'          'Then we need the bitmaps. Note that we create a monochrome bitmap here!
'          'this is a trick we use for creating a mask fast enough.
'70        hMaskBmpA = CreateBitmap(width, height, 1, 1, ByVal 0&)
'80        hTempBmpA = CreateCompatibleBitmap(dsthdc, width, height)
'          '..then we can assign the bitmaps to the DCs
'90        hMaskBmpB = SelectObject(maskDC, hMaskBmpA)
'100       hTempBmpB = SelectObject(tempDC, hTempBmpA)
'          'Now we can create a mask..First we set the background color to the
'          'transparent color then we copy the image into the monochrome bitmap.
'          'When we are done, we reset the background color of the original source.
'110       transColor = SetBkColor(srchDC, transColor)
'120       BitBlt maskDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
'130       transColor = SetBkColor(srchDC, transColor)
'          'The first we do with the mask is to MergePaint it into the destination.
'          'this will punch a WHITE hole in the background exactly were we want the
'          'graphics to be painted in.
'140       BitBlt tempDC, 0, 0, width, height, maskDC, 0, 0, vbSrcCopy
'150       BitBlt dsthdc, X, Y, width, height, tempDC, 0, 0, vbMergePaint
'          'Now we delete the transparent part of our source image. To do this
'          'we must invert the mask and MergePaint it into the source image. the
'          'transparent area will now appear as WHITE.
'160       BitBlt maskDC, 0, 0, width, height, maskDC, 0, 0, vbNotSrcCopy
'170       BitBlt tempDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
'180       BitBlt tempDC, 0, 0, width, height, maskDC, 0, 0, vbMergePaint
'          'Both target and source are clean, all we have to do is to AND them together!
'190       BitBlt dsthdc, X, Y, width, height, tempDC, 0, 0, vbSrcAnd
'          'Now all we have to do is to clean up after us and free system resources..
'
'200       DeleteObject (hMaskBmpB)
'210       DeleteObject (hTempBmpB)
'220       DeleteObject (hMaskBmpA)
'230       DeleteObject (hTempBmpA)
'240       DeleteObject (tempbmp)
'250       DeleteObject (tempbmp2)
'260       DeleteDC (tempDC)
'270       DeleteDC (maskDC)
'280       DeleteDC (srchDC)
'End Sub



Public Property Get width() As Integer
10        width = viewBMP.bmWidth
End Property

Public Property Get WidthBytes() As Integer
10        WidthBytes = viewBMP.bmWidthBytes
End Property

Public Property Get height() As Integer
10        height = viewBMP.bmHeight
End Property


Sub DrawLine(ByVal InitialX As Integer, ByVal InitialY As Integer, ByVal FinalX As Integer, ByVal FinalY As Integer, color As Long)
         ' Bresenham's line algorithm for Microsoft Visual Basic 6.0
         ' Implementation by Robert Lee <rlee0001@maine.rr.com> July, 2002 Public Domain
          
         Dim Steep As Boolean
         Dim DeltaX As Integer, DeltaY As Integer, Delta As Integer
         Dim StepX As Integer, StepY As Integer
         Dim coord As Integer
          
10       Steep = False
20       DeltaX = Abs(FinalX - InitialX)
30       If (FinalX - InitialX) > 0 Then
40           StepX = 1
50       Else
60           StepX = -1
70       End If
80       DeltaY = Abs(FinalY - InitialY)
90       If (FinalY - InitialY) > 0 Then
100          StepY = 1
110      Else
120          StepY = -1
130      End If
140      If DeltaY > DeltaX Then
150          Steep = True
160          Swap InitialX, InitialY
170          Swap DeltaX, DeltaY
180          Swap StepX, StepY
190      End If
200      Delta = (DeltaY * 2) - DeltaX
210      For coord = 0 To DeltaX - 1
220          If Steep Then
230              Call setPixelLong(InitialY, InitialX, color)
240          Else
                 'Me.PSet (InitialX, InitialY)
250              Call setPixelLong(InitialX, InitialY, color)
260          End If
270          While Delta >= 0
280              InitialY = InitialY + StepY
290              Delta = Delta - (DeltaX * 2)
300          Wend
310          InitialX = InitialX + StepX
320          Delta = Delta + (DeltaY * 2)
330      Next coord
         'Me.PSet (FinalX, FinalY)
340      Call setPixelLong(FinalX, FinalY, color)
End Sub

Private Sub Swap(ByRef a As Integer, ByRef b As Integer)
          Dim X As Integer
10        X = a
20        a = b
30        b = X
End Sub


Sub Destroy()
10        If Not bInit Then Exit Sub
            
            'clear up reference
20        CopyMemory ByVal VarPtrArray(Data), 0&, 4
            
30        Set Pic = Nothing
40        Erase Data
50        Call DeleteDC(idc)
60        Call DeleteObject(ibitmap)

End Sub
