Attribute VB_Name = "APIs"
Option Explicit


Declare Function IntersectRect Lib "user32.dll" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global Const MF_BYPOSITION = &H400&
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Global Const OF_READ = &H0&
Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Const READONLY = &H1
Const HIDDEN = &H2
Const SYSTEM = &H4
Const ARCHIVE = &H20
Const NORMAL = &H80

Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                                             (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
                                            (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Global Const COLORONCOLOR = 3
Global Const HALFTONE = 4


Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)



Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long


Declare Sub FillMemory Lib "kernel32" _
                Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
                
Declare Sub CopyMemory Lib "kernel32" _
                               Alias "RtlMoveMemory" (Destination As Any, Source As Any, _
                                                      ByVal Length As Long)
Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Declare Function lstrlenA Lib "kernel32" _
                                  (ByVal lpString As Long) As Long

Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function BringWindowToTop Lib "user32" _
   (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hWnd As Long) As Long

Type POINTAPI
    X As Long
    Y As Long
End Type


Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Private Declare Function lstrlenW Lib _
                                  "kernel32" (ByVal lpString As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" Alias _
                                               "RegisterWindowMessageA" (ByVal lpString As String) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                                 (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam _
                                                                          As Long, ByVal lParam As Long) As Long

Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101
Global Const WM_CHAR = &H102


Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long

Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Public Const PS_SOLID = 0
Public Const PS_NULL = 5



Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Public Const R2_BLACK = 1 ' 0
Public Const R2_COPYPEN = 13 ' P
Public Const R2_LAST = 16
Public Const R2_MASKNOTPEN = 3 ' DPna
Public Const R2_MASKPEN = 9 ' DPa
Public Const R2_MASKPENNOT = 5 ' PDna
Public Const R2_MERGENOTPEN = 12    ' DPno
Public Const R2_MERGEPEN = 15 ' DPo
Public Const R2_MERGEPENNOT = 14    ' PDno
Public Const R2_NOP = 11    ' D
Public Const R2_NOT = 6 ' Dn
Public Const R2_NOTCOPYPEN = 4 ' PN
Public Const R2_NOTMASKPEN = 8 ' DPan
Public Const R2_NOTMERGEPEN = 2 ' DPon
Public Const R2_NOTXORPEN = 10 ' DPxn
Public Const R2_WHITE = 16 ' 1
Public Const R2_XORPEN = 7 ' DPx

'api Global Constants
Global Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowCaret& Lib "user32" (ByVal hWnd As Long)

Public Const MSG_OPENMAP = "DCME_OPENMAP"

Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)
Public Const SM_CXSCREEN = 0        ' Width of screen
Public Const SM_CYSCREEN = 1        ' Height of screen
Public Const SM_CXFULLSCREEN = 16   ' Width of window client area
Public Const SM_CYFULLSCREEN = 17   ' Height of window client area
Public Const SM_CYMENU = 15         ' Height of menu
Public Const SM_CYCAPTION = 4       ' Height of caption or title
Public Const SM_CXFRAME = 32        ' Width of window frame
Public Const SM_CYFRAME = 33        ' Height of window frame
Public Const SM_CXHSCROLL = 21      ' Width of arrow bitmap on
'  horizontal scroll bar
Public Const SM_CYHSCROLL = 3       ' Height of arrow bitmap on
'  horizontal scroll bar
Public Const SM_CXVSCROLL = 2       ' Width of arrow bitmap on
'  vertical scroll bar
Public Const SM_CYVSCROLL = 20      ' Height of arrow bitmap on
'  vertical scroll bar
Public Const SM_CXSIZE = 30         ' Width of bitmaps in title bar
Public Const SM_CYSIZE = 31         ' Height of bitmaps in title bar
Public Const SM_CXCURSOR = 13       ' Width of cursor
Public Const SM_CYCURSOR = 14       ' Height of cursor
Public Const SM_CXBORDER = 5        ' Width of window frame that cannot
'  be sized
Public Const SM_CYBORDER = 6        ' Height of window frame that cannot
'  be sized
Public Const SM_CXDOUBLECLICK = 36  ' Width of rectangle around the
'  location of the first click. The
'  second click must occur in the
'  same rectangular location.
Public Const SM_CYDOUBLECLICK = 37  ' Height of rectangle around the
'  location of the first click. The
'  second click must occur in the
'  same rectangular location.
Public Const SM_CXDLGFRAME = 7      ' Width of dialog frame window
Public Const SM_CYDLGFRAME = 8      ' Height of dialog frame window
Public Const SM_CXICON = 11         ' Width of icon
Public Const SM_CYICON = 12         ' Height of icon
Public Const SM_CXICONSPACING = 38  ' Width of rectangles the system
' uses to position tiled icons
Public Const SM_CYICONSPACING = 39  ' Height of rectangles the system
' uses to position tiled icons
Public Const SM_CXMIN = 28          ' Minimum width of window
Public Const SM_CYMIN = 29          ' Minimum height of window
Public Const SM_CXMINTRACK = 34     ' Minimum tracking width of window
Public Const SM_CYMINTRACK = 35     ' Minimum tracking height of window
Public Const SM_CXHTHUMB = 10       ' Width of scroll box (thumb) on
'  horizontal scroll bar
Public Const SM_CYVTHUMB = 9        ' Width of scroll box (thumb) on
'  vertical scroll bar
Public Const SM_DBCSENABLED = 42    ' Returns a non-zero if the current
'  Windows version uses double-byte
'  characters, otherwise returns
'  zero
Public Const SM_DEBUG = 22          ' Returns non-zero if the Windows
'  version is a debugging version
Public Const SM_MENUDROPALIGNMENT = 40
' Alignment of pop-up menus. If zero,
'  left side is aligned with
'  corresponding left side of menu-
'  bar item. If non-zero, left side
'  is aligned with right side of
'  corresponding menu bar item
Public Const SM_MOUSEPRESENT = 19   ' Non-zero if mouse hardware is
'  installed
Public Const SM_PENWINDOWS = 41     ' Handle of Pen Windows dynamic link
'  library if Pen Windows is
'  installed
Public Const SM_SWAPBUTTON = 23     ' Non-zero if the left and right
' mouse buttons are swapped

Declare Function GetVolumeInformation& Lib "kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName _
    As String, ByVal pVolumeNameBuffer As String, ByVal _
    nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As _
    Long, ByVal lpFileSystemNameBuffer As String, ByVal _
    nFileSystemNameSize As Long)
Public Const MAX_FILENAME_LEN = 256
    





Public Function WM_DCME_OPENMAP() As Long
    Static msg As Long

    If msg = 0 Then
        msg = RegisterWindowMessage(MSG_OPENMAP)
    End If

    WM_DCME_OPENMAP = msg

End Function





'make a form normal
Sub MakeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

'make a form topmost
Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Sub RestoreWin(hWndToRestore As Long, forceMaximize As Boolean)
    On Error GoTo RestoreWindow_Error
    
   Dim currWinP As WINDOWPLACEMENT
   
  'if a window handle passed
   If hWndToRestore Then
   
     'prepare the WINDOWPLACEMENT type
     'to receive the window coordinates
     'of the specified handle
        currWinP.Length = Len(currWinP)
    
     'get the info...
        If GetWindowPlacement(hWndToRestore, currWinP) > 0 Then
        'based on the returned info,
        'determine the window state
             If currWinP.showCmd = SW_SHOWMINIMIZED Then
               'it is minimized, so restore it
                With currWinP
                   .Length = Len(currWinP)
                   .flags = 0&
                   If forceMaximize Then
                        .showCmd = SW_SHOWMAXIMIZED
                   Else
                        .showCmd = SW_RESTORE
                   End If
                End With
    
                Call SetWindowPlacement(hWndToRestore, currWinP)
             Else
               'it is on-screen, so make it visible
                Call SetForegroundWindow(hWndToRestore)
                Call BringWindowToTop(hWndToRestore)
                
             End If
        End If
    End If
   
   On Error GoTo 0
   Exit Sub
   
RestoreWindow_Error:
    HandleError Err, "RestoreWin"
End Sub



Function WinDir() As String
    Dim buf As String
    Dim ret As Long

    buf = String$(260, Chr$(0))
    ret = GetWindowsDirectory(buf, Len(buf))
    WinDir = Left$(buf, ret)
End Function

Function SysDir() As String
    Dim buf As String
    Dim ret As Long

    buf = String$(260, Chr$(0))
    ret = GetSystemDirectory(buf, Len(buf))
    SysDir = Left$(buf, ret)
End Function



Function PointerToString(lngPtr As Long) As String
'--------------------------------------------------------
'RETURNS A STRING FROM IT'S POINTER
'EXAMPLE:
'-- Generate pointer for demo purposes

'Dim l As Long
'Dim s As String
's = "THIS IS A TEST"
'l = StrPtr(s)

'--We have the pointer, call the function

'messagebox PointerToString(l)

'NOTE: THE ASSUMPTION IS THAT THE POINTER IS TO A UNICODE STRING
'IF NOT, CHANGE THE FUNCTION AS FOLLOWS (UNTESTED)
'-- Change lstrlenW to lStrLena
'-- Get rid of the * 2
'-- The replace statement should not be necessary, just return strTemp
'----------------------------------------------------------

    Dim strTemp As String
    Dim lngLen As Long


    If lngPtr Then
        lngLen = lstrlenW(lngPtr) * 2
        If lngLen Then
            strTemp = Space(lngLen)
            CopyMemory ByVal strTemp, ByVal lngPtr, lngLen
            PointerToString = replace(strTemp, Chr(0), "")
        End If
    End If
End Function


