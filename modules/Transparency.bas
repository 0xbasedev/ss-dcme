Attribute VB_Name = "Transparency"
Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function TransBltAPI Lib "msimg32.dll" Alias "TransparentBlt" _
  (ByVal hDC As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hSrcDC As Long, _
   ByVal xSrc As Long, _
   ByVal ySrc As Long, _
   ByVal nSrcWidth As Long, _
   ByVal nSrcHeight As Long, _
   ByVal crTransparent As Long) As Boolean

Private Declare Function AlphaBlend Lib "msimg32.dll" _
    (ByVal hDC As Long, _
    ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, _
    ByVal hDC As Long, _
    ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, _
    ByVal BLENDFUNCT As Long) As Long
    
Private Const AC_SRC_OVER = &H0
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type


Sub AlphaBlt(hdcDest As Long, DestX As Integer, DestY As Integer, destw As Integer, desth As Integer, hdcSrc As Long, SrcX As Integer, SrcY As Integer, srcW As Integer, srcH As Integer, alpha As Byte)
          'KPD-Team 2000
          'URL: http://www.allapi.net/
          'E-Mail: KPDTeam@Allapi.net
          Dim BF As BLENDFUNCTION, lBF As Long
          'set the parameters
10        With BF
20            .BlendOp = AC_SRC_OVER
30            .BlendFlags = 0
40            .SourceConstantAlpha = alpha
50            .AlphaFormat = 0
60        End With
          'copy the BLENDFUNCTION-structure to a Long
70        RtlMoveMemory lBF, BF, 4
          'AlphaBlend the picture from Picture1 over the picture of Picture2
80        AlphaBlend hdcDest, CLng(DestX), CLng(DestY), CLng(destw), CLng(desth), hdcSrc, CLng(SrcX), CLng(SrcY), CLng(srcW), CLng(srcH), lBF
          
End Sub

Function TransparentBlt(dsthdc As Long, X As Integer, Y As Integer, width As Integer, height As Integer, sourcehDC As Long, xSrc As Integer, ySrc As Integer, TransColor As Long) As Boolean
10        If currentWindowsVersion > os_win98 Then
20            TransparentBlt = TransBltAPI(dsthdc, X, Y, width, height, sourcehDC, xSrc, ySrc, width, height, TransColor)
30        Else
40            TransparentBltOld dsthdc, X, Y, width, height, sourcehDC, xSrc, ySrc, TransColor
50            TransparentBlt = True
60        End If
      '    If TransparentBlt = False Then Stop
End Function



Private Sub TransparentBltOld(dsthdc As Long, X As Integer, Y As Integer, width As Integer, height As Integer, sourcehDC As Long, xSrc As Integer, ySrc As Integer, TransColor As Long)
          'For some reason piclevel transparentblt fails, while picradar succeeds. I don't know why :(
          'If currentWindowsVersion <> os_win95 And currentWindowsVersion <> os_win98 Then
          '    Dim ret As Long
          '    ret = TransBltAPI(dsthdc, X, Y, width, height, sourcehDC, xSrc, ySrc, width, height, TransColor)
          '    If ret = 0 Then
          '        Dim l As Long
          '        frmGeneral.Label6.Caption = GetLastError
          '    End If
          '    Exit Sub
          'End If

          Dim maskDC As Long      'DC for the mask
          Dim tempDC As Long      'DC for temporary data
          Dim hMaskBmpA As Long    'Bitmap for mask
          Dim hTempBmpA As Long    'Bitmap for temporary data
          Dim hMaskBmpB As Long    'Bitmap for mask
          Dim hTempBmpB As Long    'Bitmap for temporary data
          Dim tempBMP As Long
          Dim tempBMP2 As Long
          Dim srchDC As Long

10        srchDC = CreateCompatibleDC(sourcehDC)
20        tempBMP = CreateCompatibleBitmap(sourcehDC, width, height)
30        tempBMP2 = SelectObject(srchDC, tempBMP)
40        BitBlt srchDC, 0, 0, width, height, sourcehDC, xSrc, ySrc, vbSrcCopy

          'First create some DC's. These are our gateways to assosiated bitmaps in RAM
50        maskDC = CreateCompatibleDC(dsthdc)
60        tempDC = CreateCompatibleDC(dsthdc)
          'Then we need the bitmaps. Note that we create a monochrome bitmap here!
          'this is a trick we use for creating a mask fast enough.
70        hMaskBmpA = CreateBitmap(width, height, 1, 1, ByVal 0&)
80        hTempBmpA = CreateCompatibleBitmap(dsthdc, width, height)
          '..then we can assign the bitmaps to the DCs
90        hMaskBmpB = SelectObject(maskDC, hMaskBmpA)
100       hTempBmpB = SelectObject(tempDC, hTempBmpA)
          'Now we can create a mask..First we set the background color to the
          'transparent color then we copy the image into the monochrome bitmap.
          'When we are done, we reset the background color of the original source.
110       TransColor = SetBkColor(srchDC, TransColor)
120       BitBlt maskDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
130       TransColor = SetBkColor(srchDC, TransColor)
          'The first we do with the mask is to MergePaint it into the destination.
          'this will punch a WHITE hole in the background exactly were we want the
          'graphics to be painted in.
140       BitBlt tempDC, 0, 0, width, height, maskDC, 0, 0, vbSrcCopy
150       BitBlt dsthdc, X, Y, width, height, tempDC, 0, 0, vbMergePaint
          'Now we delete the transparent part of our source image. To do this
          'we must invert the mask and MergePaint it into the source image. the
          'transparent area will now appear as WHITE.
160       BitBlt maskDC, 0, 0, width, height, maskDC, 0, 0, vbNotSrcCopy
170       BitBlt tempDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
180       BitBlt tempDC, 0, 0, width, height, maskDC, 0, 0, vbMergePaint
          'Both target and source are clean, all we have to do is to AND them together!
190       BitBlt dsthdc, X, Y, width, height, tempDC, 0, 0, vbSrcAnd
          'Now all we have to do is to clean up after us and free system resources..

200       DeleteObject (hMaskBmpB)
210       DeleteObject (hTempBmpB)
220       DeleteObject (hMaskBmpA)
230       DeleteObject (hTempBmpA)
240       DeleteObject (tempBMP)
250       DeleteObject (tempBMP2)
260       DeleteDC (tempDC)
270       DeleteDC (maskDC)
280       DeleteDC (srchDC)
End Sub




Sub TransAlphaBlt(dsthdc As Long, X As Integer, Y As Integer, width As Integer, height As Integer, sourcehDC As Long, xSrc As Integer, ySrc As Integer, TransColor As Long, alpha As Byte)
      'blend p2 -> p1 with transcolor tc:
      '
      '1: copy p1 --> temp
      '2: transp p2 --> temp
      '3: alphablend temp --> p1
          Dim tempDC As Long
10        tempDC = CreateCompatibleDC(dsthdc)

          Dim tempBMP As Long
          Dim tempBMP2 As Long
20        tempBMP = CreateCompatibleBitmap(sourcehDC, width, height)
30        tempBMP2 = SelectObject(tempDC, tempBMP)

          '1: copy p1 --> temp
40        BitBlt tempDC, 0, 0, width, height, dsthdc, X, Y, vbSrcCopy

        SetBkColor tempDC, GetBkColor(dsthdc)
        
          '2: tranp p2 --> temp
50        TransBltAPI tempDC, 0, 0, width, height, sourcehDC, xSrc, ySrc, width, height, TransColor

          '3: alphablend temp --> p1

          'set the parameters
          Dim BF As BLENDFUNCTION, lBF As Long
60        With BF
70            .BlendOp = AC_SRC_OVER
80            .BlendFlags = 0
90            .SourceConstantAlpha = alpha
100           .AlphaFormat = 0
110       End With
          'copy the BLENDFUNCTION-structure to a Long
120       RtlMoveMemory lBF, BF, 4

130       AlphaBlend dsthdc, X, Y, width, height, tempDC, 0, 0, width, height, lBF

          'clean up
140       DeleteObject (tempBMP)
150       DeleteObject (tempBMP2)
160       DeleteDC (tempDC)
End Sub

