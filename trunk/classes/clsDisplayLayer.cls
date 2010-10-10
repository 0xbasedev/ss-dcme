VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisplayLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
       
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, _
                                        ByVal X As Long, ByVal Y As Long, _
                                        ByVal lpString As String, _
                                        ByVal nCount As Long) As Long
                                        
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Const TEXT_TRANSPARENT = 1
Const TEXT_OPAQUE = 2
          
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
        lpRect As RECT, ByVal wFormat As Long) As Long

      Private Const DT_BOTTOM = &H8
      Private Const DT_CENTER = &H1
      Private Const DT_LEFT = &H0
      Private Const DT_RIGHT = &H2
      Private Const DT_TOP = &H0
      Private Const DT_VCENTER = &H4
      Private Const DT_WORDBREAK = &H10

' Logical Font
Const LF_FACESIZE = 32
Const LF_FULLFACESIZE = 64
Const FONT_SIZE = 12
Const NO_ERROR = 0
Const ANSI_CHARSET = 0
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_DONTCARE = 0
Const LOGPIXELSY = 90
Const TRANSPARENT = 1

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
        'Change to (255) if 8bit images can be cached
        bmiColors(15) As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long


Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long


Dim c_hDC As Long
Dim c_width As Long
Dim c_height As Long

Dim c_backColor As Long

Dim c_hBitmap As Long

Dim is_cached As Boolean
Dim cache_name As String 'Filename of the cached bitmap

Dim bmpInfo As BITMAPINFO
Dim bmpFileheader As BITMAPFILEHEADER


Public Property Get hDC() As Long
    hDC = c_hDC
End Property

Public Property Get hBitmap() As Long
    hBitmap = c_hBitmap
End Property



Public Property Get width() As Long
    width = c_width
End Property

Public Property Get height() As Long
    height = c_height
End Property

Public Property Get BackColor() As Long
    BackColor = c_backColor
End Property

Public Property Let BackColor(color As Long)
    c_backColor = color
    
    Call SetBkColor(c_hDC, color)
    
End Property



Public Property Let FontSize(sf_height As Long)
    Dim hFont As Long, hOldFont As Long
    Dim lf As LOGFONT
    
'    With lf
    lf.lfHeight = sf_height
'    End With
    
    hFont = CreateFontIndirectA(lf)
    
    hOldFont = SelectObject(c_hDC, hFont)
    
    
    DeleteObject hOldFont
End Property


Public Property Let TextBackMode(bkmode As Long)
    Call SetBkMode(c_hDC, bkmode)
End Property

Public Property Let TextColor(color As Long)
    Call SetTextColor(c_hDC, color)
End Property

Friend Sub PrintText(Text As String, X As Long, Y As Long)
    Call SetBkMode(c_hDC, TEXT_TRANSPARENT)
    TextOut c_hDC, X, Y, Text, Len(Text)
End Sub



Friend Sub ColorReplace(oldcolor As Long, newcolor As Long)
    
    If bmpInfo.bmiHeader.biBitCount <= 8 Then Exit Sub
    
    
    
    
    'Create a new hDC, with the new color as background
    Dim tmp_bmp As Long, oldbmp As Long
    Dim tmphDC As Long, oldhDC As Long
    
    oldhDC = c_hDC
    
    Dim dczero As Long
    dczero = GetDC(0)
    
    'Get rid of the old hDC, use the new one
    c_hDC = CreateCompatibleDC(dczero)
    tmp_bmp = CreateCompatibleBitmap(dczero, width, height)
    oldbmp = SelectObject(c_hDC, tmp_bmp)
    
    Call ReleaseDC(0&, dczero)
    
    'Fill the new hdc with newcolor
    DrawFilledRectangle c_hDC, 0, 0, CInt(width), CInt(height), newcolor
    
    'TransBLT the old image on the new hdc, using oldcolor as transparency key
    TransparentBlt c_hDC, 0, 0, width, height, oldhDC, 0, 0, oldcolor
    
    DeleteDC oldhDC
    DeleteObject hBitmap
    DeleteObject oldbmp
'        DeleteObject tmpsel
    
    c_hBitmap = tmp_bmp

End Sub







'These are simply shortcuts when you want to copy from one layer to another with the same coordinates

Friend Sub BitBltToLayerFull(ByRef destLayer As clsDisplayLayer, dwRop As Long)
    BitBlt destLayer.hDC, 0, 0, c_width, c_height, c_hDC, 0, 0, dwRop
End Sub

Friend Sub TransparentBltToLayerFull(ByRef destLayer As clsDisplayLayer, TransColor As Long)
    TransparentBlt destLayer.hDC, 0, 0, CInt(c_width), CInt(c_height), c_hDC, 0, 0, TransColor
End Sub

Friend Sub AlphaBlendToLayerFull(ByRef destLayer As clsDisplayLayer, alpha As Byte)
    AlphaBlt destLayer.hDC, 0, 0, CInt(c_width), CInt(c_height), c_hDC, 0, 0, CInt(c_width), CInt(c_height), alpha
End Sub

Friend Sub AlphaTransToLayerFull(ByRef destLayer As clsDisplayLayer, TransColor As Long, alpha As Byte)
'    If left < 0 Then left = 0
'    If top < 0 Then top = 0
'    If right > width Then right = width
'    If bottom > height Then bottom = height
    
    TransAlphaBlt destLayer.hDC, 0, 0, CInt(c_width), CInt(c_height), c_hDC, 0, 0, TransColor, alpha
End Sub


Friend Sub BitBltToLayer(ByRef destLayer As clsDisplayLayer, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, dwRop As Long)
    If Left < 0 Then Left = 0
    If Top < 0 Then Top = 0
    If Right >= width Then Right = width - 1
    If Bottom >= height Then Bottom = height - 1
    
    If BitBlt(destLayer.hDC, Left, Top, Right - Left + 1, Bottom - Top + 1, c_hDC, Left, Top, dwRop) = 0 Then MessageBox "BitBlt failed!"
End Sub

Friend Sub TransparentBltToLayer(ByRef destLayer As clsDisplayLayer, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, TransColor As Long)
    If Left < 0 Then Left = 0
    If Top < 0 Then Top = 0
    If Right >= width Then Right = width - 1
    If Bottom >= height Then Bottom = height - 1
    
    If Right > Left And Bottom > Top Then
        If TransparentBlt(destLayer.hDC, Left, Top, Right - Left + 1, Bottom - Top + 1, c_hDC, Left, Top, TransColor) = False Then MessageBox "TransparentBlt failed!"
    End If
End Sub

Friend Sub AlphaBlendToLayer(ByRef destLayer As clsDisplayLayer, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, alpha As Byte)
'    If left < 0 Then left = 0
'    If top < 0 Then top = 0
'    If right > width Then right = width
'    If bottom > height Then bottom = height
    
    AlphaBlt destLayer.hDC, Left, Top, Right - Left + 1, Bottom - Top + 1, c_hDC, Left, Top, Right - Left + 1, Right - Left + 1, alpha
End Sub

Friend Sub AlphaTransToLayer(ByRef destLayer As clsDisplayLayer, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, TransColor As Long, alpha As Byte)
'    If left < 0 Then left = 0
'    If top < 0 Then top = 0
'    If right > width Then right = width
'    If bottom > height Then bottom = height
    
    TransAlphaBlt destLayer.hDC, Left, Top, Right - Left + 1, Bottom - Top + 1, c_hDC, Left, Top, TransColor, alpha
End Sub


Sub LoadImageFile(path As String, ByRef tmpPic As PictureBox, ByRef p_width As Integer, ByRef p_height As Integer)
    tmpPic.AutoSize = True
    
    Call LoadPic(tmpPic, path)
    
    p_width = tmpPic.ScaleWidth
    p_height = tmpPic.ScaleHeight
    
    Call Resize(p_width, p_height, False)
    
    BitBlt c_hDC, 0, 0, p_width, p_height, tmpPic.hDC, 0, 0, vbSrcCopy
    
    tmpPic.AutoSize = False
    
    tmpPic.Cls
    
End Sub



Friend Sub MoveBitmap(dx As Integer, dy As Integer)
    If dx <> 0 Or dy <> 0 Then
        If Abs(dx) < c_width And Abs(dy) < c_height Then
            Dim SrcX As Long, SrcY As Long
            Dim DestX As Long, DestY As Long
            Dim mvWidth As Long, mvHeight As Long

            
            If dx > 0 Then
                mvWidth = c_width - dx
                SrcX = 0
                DestX = dx
            Else
                mvWidth = c_width + dx
                SrcX = -dx
                DestX = 0
            End If

            If dy > 0 Then
                mvHeight = c_height - dy
                SrcY = 0
                DestY = dy
            Else
                mvHeight = c_height + dy
                SrcY = -dy
                DestY = 0
            End If
            
            BitBlt c_hDC, DestX, DestY, mvWidth, mvHeight, c_hDC, SrcX, SrcY, vbSrcCopy
            
        End If
    End If
End Sub





Sub EraseArea(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer)
    If Not is_cached Then
        DrawFilledRectangle c_hDC, Left, Top, Right, Bottom, c_backColor
    End If
End Sub

Sub Cls()
    If c_backColor = vbBlack Then
      Dim oldbmp As Long
      
        DeleteObject hBitmap

        c_hBitmap = CreateCompatibleBitmap(c_hDC, c_width, c_height)
        oldbmp = SelectObject(c_hDC, c_hBitmap)

      DeleteObject oldbmp
    Else
        Call EraseArea(0, 0, CInt(c_width), CInt(c_height))
    End If
End Sub

Sub Resize(width As Integer, height As Integer, keepPrevious As Boolean)
    Dim tmp_bmp As Long, oldbmp As Long, tmphDC As Long
    
    If is_cached Then
        Err.Raise 20010, "clsDisplayLayer.Resize", "Attempted to resize a cached layer"
        Exit Sub
    End If
  
    If keepPrevious Then
        
        
        Dim dczero As Long
        dczero = GetDC(0)
    
        tmphDC = CreateCompatibleDC(dczero)
        tmp_bmp = CreateCompatibleBitmap(dczero, width, height)
        oldbmp = SelectObject(tmphDC, tmp_bmp)
        
        DeleteObject oldbmp
        Call ReleaseDC(0&, dczero)
        
        BitBlt tmphDC, 0, 0, longMinimum(CLng(width), c_width), longMinimum(CLng(height), c_height), c_hDC, 0, 0, vbSrcCopy
        
        
        DeleteDC c_hDC
        DeleteObject hBitmap

'        DeleteObject tmpsel
        
        c_hDC = tmphDC
        c_hBitmap = tmp_bmp

    Else
        DeleteObject hBitmap
        
        c_hBitmap = CreateCompatibleBitmap(c_hDC, width, height)
        oldbmp = SelectObject(c_hDC, hBitmap)
      DeleteObject oldbmp
      
    End If
    
    c_width = width
    c_height = height


    'Update headers
'    With bmpInfo.bmiHeader
'        '.biSize = 40
'        .biWidth = c_width
'        .biHeight = c_height
'        '.biPlanes = 1
'        '.biBitCount = 24
'        '.biCompression = 0& 'BI_RGB
'        .biSizeImage = c_height * ((c_width * .biBitCount \ 8 + 3) And &HFFFFFFFC)
''        .biSizeImage = c_height * ((c_width * 3 + 3) And &HFFFFFFFC)
'
'        bmpFileheader.bfSize = .biSizeImage
'    End With
    
    'Make it generate the header
    Dim tmpheader As BITMAPINFO
    tmpheader.bmiHeader.biSize = 40
    
    GetDIBits c_hDC, hBitmap, 0, c_height, ByVal 0&, tmpheader, 0
    
    bmpInfo = tmpheader
    bmpFileheader.bfSize = bmpInfo.bmiHeader.biSizeImage
End Sub


Sub SetBitCount(bcount As Integer)

    
    Dim oldbmp As Long
    oldbmp = hBitmap
    
    c_hBitmap = CreateBitmap(c_width, c_height, 1, 1, ByVal 0&)
    
   
    'Should return oldbmp
    SelectObject c_hDC, hBitmap
    

    
    DeleteObject oldbmp
    
    
    'Make it generate the header
    Dim tmpheader As BITMAPINFO
    tmpheader.bmiHeader.biSize = 40
    
    GetDIBits c_hDC, hBitmap, 0, c_height, ByVal 0&, tmpheader, 0
    
    bmpInfo = tmpheader
    bmpFileheader.bfSize = bmpInfo.bmiHeader.biSizeImage
    
'    With bmpInfo.bmiHeader
'        .biBitCount = bcount
'        If tmpheader.bmiHeader.biSizeImage <> 0 Then
'            .biSizeImage = tmpheader.bmiHeader.biSizeImage
'        Else
'            .biSizeImage = c_height * ((c_width * bcount \ 8 + 3) And &HFFFFFFFC)
'        End If
'
'
'    End With
End Sub

Private Sub Class_Initialize()
    c_width = 1
    c_height = 1

    Dim dczero As Long, oldbmp As Long
    dczero = GetDC(0)
    
    c_hDC = CreateCompatibleDC(dczero)
    SetBkColor c_hDC, c_backColor
    
    c_hBitmap = CreateCompatibleBitmap(dczero, c_width, c_height)
    oldbmp = SelectObject(c_hDC, hBitmap)

    DeleteObject oldbmp
  
    Call ReleaseDC(0&, dczero)


    With bmpInfo.bmiHeader
        .biSize = 40
        .biWidth = c_width
        .biHeight = c_height
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = 0& 'BI_RGB
        .biSizeImage = c_height * ((c_width * 3 + 3) And &HFFFFFFFC)
    End With
    
    With bmpFileheader
        .bfType = &H4D42
        .bfSize = bmpInfo.bmiHeader.biSizeImage
        .bfOffBits = 54
    End With
End Sub

Private Sub Class_Terminate()

    DeleteObject hBitmap
    DeleteDC c_hDC
    c_hDC = 0
    If is_cached Then
        DeleteFile cache_name
    End If
End Sub

Public Property Get cached() As Boolean
    cached = is_cached
End Property

Public Property Let cached(cache As Boolean)
    If is_cached <> cache Then
        is_cached = cache
        
        If is_cached Then
            Call CacheData
        Else
            Call UncacheData
        End If
    End If
End Property

Private Sub CacheData()
    Dim i As Integer
    Dim tmpname As String
    
    CreateDir Directory_Cache

    'Find a filename for our data
    tmpname = Directory_Cache & "\cache" & GetTickCount
    cache_name = tmpname
    
    While FileExists(cache_name & ".bmp")
        i = i + 1
        cache_name = tmpname & "_" & i
    Wend
    cache_name = cache_name & ".bmp"
    
    Call SaveToFile(cache_name, False)
    
    'Clean up
    DeleteDC c_hDC
    DeleteObject hBitmap
    
    c_hDC = 0
End Sub

Private Sub UncacheData()
    Call LoadFromFile(cache_name)
End Sub



Public Function GetBitmapLineLenght() As Long
    'Returns the byte lenght of a single line of bitmap data
    GetBitmapLineLenght = bmpInfo.bmiHeader.biSizeImage \ c_height
End Function

Public Function GetBitmapData(ByRef b() As Byte) As Long

    ReDim b(bmpInfo.bmiHeader.biSizeImage - 1)

    'Grab the bitmap data
    GetBitmapData = GetDIBits(c_hDC, c_hBitmap, 0, c_height, b(0), bmpInfo, 0)

End Function

Public Function GetBitmapDataLine(ByRef b() As Byte, start As Long) As Long
    'Grabs a single line of bitmap data
    
    ReDim b((bmpInfo.bmiHeader.biSizeImage \ c_height) - 1)

    'Grab the bitmap data
    GetBitmapDataLine = GetDIBits(c_hDC, c_hBitmap, c_height - 1 - start, 1, b(0), bmpInfo, 0)
End Function



Sub SaveToFile(path As String, askoverwrite As Boolean)
    Dim f As Integer, i As Integer
    
    If FileExists(path) Then
        If askoverwrite Then
            If Not CheckOverwrite(path) Then
                Exit Sub
            End If
        Else
            Call DeleteFile(path)
        End If
    End If

    'Dump the bitmap data in a file
    Dim cache_data() As Byte
    ReDim cache_data(bmpInfo.bmiHeader.biSizeImage - 1)
    
    
    'Grab the bitmap data

    
    f = FreeFile
    Open path For Binary As #f
    
    Put #f, , bmpFileheader
    Put #f, , bmpInfo.bmiHeader
    
    If bmpInfo.bmiHeader.biBitCount <= 8 Then
        Dim tmpbmpinfo As BITMAPINFO
        tmpbmpinfo = bmpInfo
    
        Dim ncolors As Long
        ncolors = bmpInfo.bmiHeader.biClrUsed
        
        GetDIBits c_hDC, hBitmap, 0, c_height, cache_data(0), tmpbmpinfo, 0
        
        For i = 0 To ncolors - 1
            Put #f, , tmpbmpinfo.bmiColors(i)
        Next
'        CopyMemory ByVal VarPtr(ctable(0)), ByVal VarPtr(bmpInfo.bmiColors(0)),
    
'        ReDim bmpInfo.bmiColors((2 ^ bmpInfo.bmiHeader.biBitCount) - 1)
    Else
        GetDIBits c_hDC, hBitmap, 0, c_height, cache_data(0), bmpInfo, 0
    End If
    
    
    
    Put #f, , cache_data
    
    Close #f
End Sub

Sub LoadFromFile(path As String)
    If FileExists(path) Then
    
        'Clean up
        DeleteDC c_hDC
        DeleteObject hBitmap
    
        c_hDC = 0
    
        Dim dczero As Long, oldbmp As Long
        Dim cache_data() As Byte
        Dim f As Integer
        Dim tmp_fileheader As BITMAPFILEHEADER
        Dim tmp_bmpinfo As BITMAPINFO
        
        

    

        f = FreeFile
        Open path For Binary As #f
        
        Get #f, , tmp_fileheader
        Get #f, , tmp_bmpinfo.bmiHeader
        
        c_width = tmp_bmpinfo.bmiHeader.biWidth
        c_height = tmp_bmpinfo.bmiHeader.biHeight
        
        dczero = GetDC(0)
    
        c_hDC = CreateCompatibleDC(dczero)
        SetBkColor c_hDC, c_backColor
        c_hBitmap = CreateCompatibleBitmap(dczero, c_width, c_height)
        oldbmp = SelectObject(c_hDC, hBitmap)
    
        DeleteObject oldbmp
      
        Call ReleaseDC(0&, dczero)
        
        
        
        If tmp_bmpinfo.bmiHeader.biClrUsed > 0 And tmp_bmpinfo.bmiHeader.biBitCount <= 8 Then
        
            Dim ctable() As RGBQUAD
            ReDim ctable(tmp_bmpinfo.bmiHeader.biClrUsed - 1)
            
            Get #f, , ctable
        
            CopyMemory ByVal VarPtr(tmp_bmpinfo.bmiColors(0)), ByVal VarPtr(ctable(0)), bmpInfo.bmiHeader.biClrUsed * 4
        End If
        
'        If tmp_bmpinfo.bmiHeader.biSizeImage <> bmpInfo.bmiHeader.biSizeImage Then
'            'Mismatch!
'            Close #f
'            Err.Raise 20012, "clsDisplayLayer.UncacheData", "Cache file '" & cache_name & "' does not have the correct size. Expected " & bmpInfo.bmiHeader.biSizeImage & ", found " & tmp_bmpinfo.bmiHeader.biSizeImage
'            Exit Sub
'        End If
        
        bmpInfo = tmp_bmpinfo
        
        
        ReDim cache_data(bmpInfo.bmiHeader.biSizeImage - 1)
        Get #f, , cache_data
        
        Close #f
        
        SetDIBits c_hDC, hBitmap, 0, c_height, cache_data(0), bmpInfo, 0
        
        DeleteFile path
    Else
        Err.Raise 20011, "clsDisplayLayer.UncacheData", "Cache file '" & cache_name & "' not found."
    End If
End Sub

