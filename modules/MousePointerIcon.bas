Attribute VB_Name = "MousePointerIcon"
Option Explicit

' used to convert a memory handle to a stdPicture object
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
                                                  (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
                                                   IPic As IPicture) As Long
Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

' used to retrieve and set information for a icon and/or cursor
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long

' drawing functions
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32.dll" (ByVal hDC As Long, _
                                                                 ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "GDI32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "GDI32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32.dll" (ByVal hDC As Long, _
                                                       ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, _
                                                    ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "GDI32.dll" (ByVal hDestDC As Long, _
                                                 ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "GDI32.dll" (ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateSolidBrush Lib "GDI32.dll" (ByVal crColor As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' retrieve information about a bitmap
Private Declare Function GetGDIObject Lib "GDI32.dll" Alias "GetObjectA" _
                                      (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Function CreateIcon32x32(hIcon16x16 As Long) As Long

    Dim hIcon As Long
    Dim icoInfo As ICONINFO, bmpInfo As BITMAP
    Dim hBmp As Long, hbmpOld As Long, hBmpSrc As Long
    Dim tmpDC As Long, tmpDCsrc As Long
    Dim tRect As RECT
    Dim hBrush As Long

    GetIconInfo hIcon16x16, icoInfo

    If icoInfo.hbmColor = 0 Then
        If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
        'exit here. Black & white icons need a bit more work, and
        'if you are converting them, let you do the leg work.
        ' Feeling helpful, but not too motivated, sorry.
        Exit Function
    End If
    If GetGDIObject(icoInfo.hbmColor, Len(bmpInfo), bmpInfo) = 0 Then Exit Function

    tRect.Right = 32
    tRect.Bottom = 32

  Dim dczero As Long
  
  dczero = GetDC(0&)
  
    ' create 2 temporary DCs. Must destroy these later
    tmpDC = CreateCompatibleDC(dczero)
    tmpDCsrc = CreateCompatibleDC(dczero)

    ' create a temp bitmap to be the new icon & select it into the DC
    hBmp = CreateCompatibleBitmap(dczero, 32, 32)

    hbmpOld = SelectObject(tmpDC, hBmp)
    ' select the icon into the other DC
    hBmpSrc = SelectObject(tmpDCsrc, icoInfo.hbmColor)
    ' fill the new bitmap with black
    hBrush = CreateSolidBrush(vbBlack)
    FillRect tmpDC, tRect, hBrush
    DeleteObject hBrush
    ' now copy of the icon into new bitmap, centering it on the way
    BitBlt tmpDC, 0, 0, _
           bmpInfo.bmWidth, bmpInfo.bmHeight, tmpDCsrc, 0, 0, vbSrcCopy
    '(32 - bmpInfo.bmWidth) \ 2, (32 - bmpInfo.bmHeight) \ 2

    ' kill the icon's bitmap (doesn't delete the icon)
    DeleteObject SelectObject(tmpDCsrc, hBmpSrc)
    ' make the new bitmap the icon's bitmap
    icoInfo.hbmColor = SelectObject(tmpDC, hbmpOld)

    ' the same exact remarks above apply to the mask below
    hBmp = CreateBitmap(32, 32, 1, 1, ByVal 0&)
    hbmpOld = SelectObject(tmpDC, hBmp)
    hBmpSrc = SelectObject(tmpDCsrc, icoInfo.hbmMask)
    hBrush = CreateSolidBrush(vbWhite)
    FillRect tmpDC, tRect, hBrush
    DeleteObject hBrush
    BitBlt tmpDC, 0, 0, _
           bmpInfo.bmWidth, bmpInfo.bmHeight, tmpDCsrc, 0, 0, vbSrcCopy
    '(32 - bmpInfo.bmWidth) \ 2, (32 - bmpInfo.bmHeight) \ 2
    DeleteObject SelectObject(tmpDCsrc, hBmpSrc)
    icoInfo.hbmMask = SelectObject(tmpDC, hbmpOld)

    ' destroy the DCs
    DeleteDC tmpDC
    DeleteDC tmpDCsrc

    ' set the cursor's hot spot. Here it will be top left of visible icon
    ' Adjust this as needed if your cursor needs to be at one of the other corners
    icoInfo.xHotspot = (32 - bmpInfo.bmWidth) \ 2
    icoInfo.yHotspot = (32 - bmpInfo.bmHeight) \ 2
    icoInfo.fIcon = 0   ' identifies the icon as a cursor vs icon

    ' create the icon & delete the bitmaps
    hIcon = CreateIconIndirect(icoInfo)

    DeleteObject icoInfo.hbmColor
    DeleteObject icoInfo.hbmMask

    ReleaseDC 0&, dczero
    ' return the result
    CreateIcon32x32 = hIcon


End Function

Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As IPicture

    On Error GoTo ExitRoutine

    Dim Pic As PICTDESC
    Dim GUID(0 To 3) As Long

    ' initialize the PictDesc structure
    Pic.cbSize = Len(Pic)
    If isBitmap Then Pic.pictType = vbPicTypeBitmap Else Pic.pictType = vbPicTypeIcon
    Pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect Pic, GUID(0), True, HandleToPicture

ExitRoutine:
End Function

