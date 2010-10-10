Attribute VB_Name = "mdlFloydSteinberg"
'Floyd Steinberg Filter Module

Option Explicit
Public WorkFilterG As Boolean
Public Enum iFilterFS
    iFLOYDSTEINBERG = 21
End Enum
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'-------------------------------------------Private var
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0&

Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup
Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure
Private mProgress As Long         '% filter progress
Private Speed(0 To 765) As Long   'Speed up values

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER   '40 bytes
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

Public Sub FloydSteinberg(Pic As PictureBox)
      'Run Filter function
10        Call FilterBW(Pic.Image, 15, mProgress)
          'Make graphics persistent
20        Pic.Picture = Pic.Image
          'Check if picture loaded...
          '  If pic.Picture <> 0 Then
          '    'Copy main picturebox to back buffer
          '    pic.PaintPicture pic.Picture, 0, 0
          '  End If
          '  'Copy back buffer to main picturebox
          '  pic.PaintPicture pic.Picture, 0, 0
End Sub
Public Sub FilterBW(ByVal Pic As Long, ByVal Factor As Long, ByRef pProgress As Long)
          Dim hdcNew As Long
          Dim oldhand As Long
          Dim ret As Long
          Dim BytesPerScanLine As Long
          Dim PadBytesPerScanLine As Long

10        If WorkFilterG = True Then Exit Sub
20        WorkFilterG = True
          'get data buffer
30        Call GetObject(Pic, Len(PicInfo), PicInfo)
40        hdcNew = CreateCompatibleDC(0&)
50        oldhand = SelectObject(hdcNew, Pic)
60        With DIBInfo.bmiHeader
70            .biSize = 40
80            .biWidth = PicInfo.bmWidth
90            .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
100           .biPlanes = 1
110           .biBitCount = 32
120           .biCompression = BI_RGB
130           BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
140           PadBytesPerScanLine = _
              BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
150           .biSizeImage = BytesPerScanLine * Abs(.biHeight)
160       End With
          'redimension  (BGR+pad,x,y)
170       ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
180       ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
          'get bytes
190       ret = GetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
200       ret = GetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, bDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
          'do it
210       Call FloydSteinbergBW(pProgress, Factor)

          'copy bytes to device
220       ret = SetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
230       SelectObject hdcNew, oldhand
240       DeleteDC hdcNew
250       ReDim iDATA(1 To 4, 1 To 2, 1 To 2) As Byte
260       ReDim bDATA(1 To 4, 1 To 2, 1 To 2) As Byte
270       WorkFilterG = False
280       Exit Sub
End Sub
Private Sub FloydSteinbergBW(ByRef pProgress As Long, ByVal PalWeight As Long)
          Dim X As Long, Y As Long
          Dim r As Long, g As Long, b As Long
          Dim Erro As Long
          Dim VecErro() As Long
          Dim nCol As Long, mCol As Long
          Dim PartErr(1 To 4, -255 To 255) As Long

10        For X = 0 To 765
20            Speed(X) = X \ 3
30        Next X
40        For X = -255 To 255
50            PartErr(1, X) = (7 * X) \ 16
60            PartErr(2, X) = (3 * X) \ 16
70            PartErr(3, X) = (5 * X) \ 16
80            PartErr(4, X) = (1 * X) \ 16
90        Next X
100       Erro = 0
110       ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
120       For X = 1 To PicInfo.bmWidth
130           VecErro(1, X) = 0
140           VecErro(2, X) = 0
150       Next X
160       pProgress = 2
170       DoEvents
180       For Y = 1 To PicInfo.bmHeight
190           For X = 1 To PicInfo.bmWidth
200               b = CLng(bDATA(1, X, Y))
210               g = CLng(bDATA(2, X, Y))
220               r = CLng(bDATA(3, X, Y))
230               b = Speed(r + g + b)
240               mCol = mCol + b
250               nCol = nCol + 1
260           Next X
270       Next Y
280       mCol = mCol \ nCol
290       pProgress = 10
300       DoEvents
310       For Y = 1 To PicInfo.bmHeight
320           For X = 1 To PicInfo.bmWidth
330               b = CLng(bDATA(1, X, Y))
340               g = CLng(bDATA(2, X, Y))
350               r = CLng(bDATA(3, X, Y))
360               b = Speed(r + g + b)
370               b = b + (VecErro(1, X) * 10) \ PalWeight
380               If b < 0 Then b = 0
390               If b > 255 Then b = 255
400               If b < mCol Then nCol = 0 Else nCol = 255
410               iDATA(1, X, Y) = nCol
420               iDATA(2, X, Y) = nCol
430               iDATA(3, X, Y) = nCol
440               Erro = b - nCol
450               If X < PicInfo.bmWidth Then VecErro(1, X + 1) = VecErro(1, X + 1) + PartErr(1, Erro)
460               If Y < PicInfo.bmHeight Then
470                   If X > 1 Then VecErro(2, X - 1) = VecErro(2, X - 1) + PartErr(2, Erro)
480                   VecErro(2, X) = VecErro(2, X) + PartErr(3, Erro)
490                   If X < PicInfo.bmWidth Then VecErro(2, X + 1) = VecErro(2, X + 1) + PartErr(4, Erro)
500               End If
510           Next X
520           For X = 1 To PicInfo.bmWidth
530               VecErro(1, X) = VecErro(2, X)
540               VecErro(2, X) = 0
550           Next X
560           mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
570           pProgress = mProgress
580           DoEvents
590       Next Y
600       pProgress = 100
610       DoEvents
End Sub


