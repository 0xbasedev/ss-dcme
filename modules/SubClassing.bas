Attribute VB_Name = "SubClassing"
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

'For SetWindowLong index
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)


'16CC0000
'16CE0000

'WS_DLGFRAME
'WS_BORDER

Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW



Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function setParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long



Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   
Private Declare Function CreateWindowEx _
    Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long


Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const WM_DESTROY = &H2
Public Const WM_CLOSE = &H10


Private lpPrevWndProc As Long

Public otherinstance As Long

Public processingMessage As Boolean
Public startingGeneral As Boolean

Dim messageQueue() As Long
Dim messageQueueCount As Long

Global Const MSG_WINDOW_TITLE = "DCME_MESSAGE_HANDLER"

Private lpPrevGeneralWndProc As Long

Sub HookWnd(hWnd As Long)
10        If Not bDEBUG Then
20            lpPrevGeneralWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf GeneralWindowProc)
30        End If
End Sub

Sub UnHookWnd(hWnd As Long)
10        If Not bDEBUG Then
20            If lpPrevGeneralWndProc Then
30                SetWindowLong hWnd, GWL_WNDPROC, lpPrevGeneralWndProc
40                lpPrevGeneralWndProc = 0
50            End If
60        End If
End Sub

Function GeneralWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          Dim zDelta As Long
          Dim curpos As POINTAPI
10        Select Case uMsg
          Case WM_DESTROY
20            UnHookWnd hWnd

30        Case WM_MOUSEWHEEL

40            Call GetCursorPos(curpos)

50            zDelta = (wParam \ 65535) Mod 65535


60            If zDelta < 0 Then
70                If frmGeneral.GetShift Then
80                    If curtool < T_lvz Then
90                        frmGeneral.SetCurrentTool (curtool + 1)
100                   End If
110               Else
120                   Call frmGeneral.ExecuteZoomFocus(True, curpos)
130               End If
140           Else
150               If frmGeneral.GetShift Then
160                   If curtool > T_magnifier Then
170                       frmGeneral.SetCurrentTool (curtool - 1)
180                   End If
190               Else
200                   Call frmGeneral.ExecuteZoomFocus(False, curpos)
210               End If
220           End If
230       Case Else
240           GeneralWindowProc = CallWindowProc(lpPrevGeneralWndProc, hWnd, uMsg, wParam, lParam)
250       End Select
End Function

Function HookMessageHandler() As Long
10        If bDEBUG Then Exit Function
          
20        otherinstance = CreateWindowEx(0, "STATIC", MSG_WINDOW_TITLE, 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)

          
30        lpPrevWndProc = SetWindowLong(otherinstance, GWL_WNDPROC, AddressOf WindowProc)

40        ReDim messageQueue(10)
50        messageQueueCount = 0
              
60        HookMessageHandler = otherinstance
End Function

Sub UnHookMessageHandler()
10        If bDEBUG Then Exit Sub

20        If lpPrevWndProc Then
30            SetWindowLong otherinstance, GWL_WNDPROC, lpPrevWndProc
40            lpPrevWndProc = 0
50        End If

60        If otherinstance Then
70          Call DestroyWindow(otherinstance)
80        End If
End Sub


'Private Sub Hook(ByVal gHW As Long)
'10        If bDEBUG Then Exit Sub
'20        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
'
'        ReDim messageQueue(10)
'        messageQueueCount = 0
'End Sub
'
'Private Sub UnHook(ByVal gHW As Long)
'10        If bDEBUG Then Exit Sub
'20        If lpPrevWndProc Then
'30            SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
'40            lpPrevWndProc = 0
'50        End If
'End Sub

Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          Dim zDelta As Long
          Dim curpos As POINTAPI
10        Select Case uMsg
          Case WM_DESTROY
20            UnHookMessageHandler

30        Case WM_MOUSEWHEEL

40            Call GetCursorPos(curpos)

50            zDelta = (wParam \ 65535) Mod 65535


60            If zDelta < 0 Then
70                If frmGeneral.GetShift Then
80                    If curtool < T_lvz Then
90                        frmGeneral.SetCurrentTool (curtool + 1)
100                   End If
110               Else
120                   Call frmGeneral.ExecuteZoomFocus(True, curpos)
130               End If
140           Else
150               If frmGeneral.GetShift Then
160                   If curtool > T_magnifier Then
170                       frmGeneral.SetCurrentTool (curtool - 1)
180                   End If
190               Else
200                   Call frmGeneral.ExecuteZoomFocus(False, curpos)
210               End If
220           End If

230       Case WM_DCME_OPENMAP    'user-defined
              'wparam contains the pointer to the string we need
              'Dim path As String
              'path = PointerToString(wParam)
240             AddDebug " +++ Received message '" & WM_DCME_OPENMAP & "', wParam=" & wParam & " lParam=" & lParam, True
                
250             Call ProcessOpenMap(wParam, lParam)

              
              
              'Call PostMessage(lParam, &H10, 0, 0)
260       Case Else
270           WindowProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
280       End Select
End Function


Private Sub AddQueue(wParam As Long)
10        If UBound(messageQueue) >= messageQueueCount Then
20            ReDim Preserve messageQueue(UBound(messageQueue) + 10)
30        End If
          
40        messageQueue(messageQueueCount) = wParam
          
50        messageQueueCount = messageQueueCount + 1
End Sub


Sub ExecuteQueuedMessages()
10        If messageQueueCount > 0 Then

20            AddDebug "Processing queued message #1 / " & messageQueueCount & ": " & messageQueue(0)
              
              
              Dim toProcess As Long
30            toProcess = messageQueue(0)
              
              'Remove it from queue
              Dim i As Long
40            For i = 1 To messageQueueCount - 1
50              messageQueue(i - 1) = messageQueue(i)
60            Next
70            messageQueueCount = messageQueueCount - 1
      '
80            Call ProcessOpenMap(toProcess, 0)
      '
90        End If
End Sub

Private Sub ProcessOpenMap(wParam As Long, lParam As Long)
10            If processingMessage Or startingGeneral Then
20              AddDebug "Message still being processed. Message queued #" & messageQueueCount & ": " & wParam
30              Call AddQueue(wParam)
40              Exit Sub
50            End If
              
60            On Error Resume Next
              
              


              Dim paths As String
              
      '              If bDEBUG Then
      '                paths = "C:\Jeux\Continuum\DCME\paths" & wParam
      '              Else
70            paths = App.path & "\paths" & wParam
      '              End If
              
              
              
80            If Not FileExists(paths) Then
90                Exit Sub
100           End If

110           processingMessage = True

              
              Dim f As Integer
120           f = FreeFile

130           Open paths For Input As #f
              Dim str As String
              
140           Line Input #f, str
              
150           AddDebug " +++ Read '" & str & "' in '" & paths & "'", True
              
160           Close #f
              
170           AddDebug " +++ Deleting '" & paths & "'", True
180           DeleteFile paths
              
            
      '        On Local Error Resume Next
              
              'bring back DCME in front (this is also needed to avoid a crash caused when DCME is
              '                          minimized while a map opens)
      '        screenUpdating = False
              
190           Call frmGeneral.RestoreWindow
200           frmGeneral.setfocus
              
              Dim mapidx As Integer
              
210           mapidx = frmGeneral.GetIndexOfMap(str)
              
220           If mapidx <> -1 Then
230             AddDebug frmGeneral.hWnd & " +++ Map '" & str & "' already opened", True
                'map already opened
240             Call frmGeneral.ActivateMap(mapidx)
250           Else
260             AddDebug frmGeneral.hWnd & " +++ Opening '" & str & "'", True
270             Call frmGeneral.OpenMap(str)
280             AddDebug frmGeneral.hWnd & " +++ OPENED '" & str & "'", True
        '        screenUpdating = True
290           End If
              
300           processingMessage = False
              
310           Call ExecuteQueuedMessages
              
End Sub

'Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
'          Dim sSave As String, ret As Long
'10        ret = GetWindowTextLength(hWnd)
'20        sSave = Space(ret)
'30        GetWindowText hWnd, sSave, ret + 1
'          'continue enumeration
'
''          AddDebug "Found: " & sSave
'40        If InStr(sSave, "Drake Continuum Map Editor") > 0 Then
'              Dim strclass As String
'50            strclass = String$(50, Chr(0))
'60            Call GetClassName(hWnd, strclass, 50)
'
''              AddDebug " Class: " & strclass
'
'70            If InStr(strclass, "ThunderRT6MDIForm") > 0 Or InStr(strclass, "ThunderMDIForm") > 0 Then
'80                otherinstance = hWnd
'90                Exit Function
'100           End If
'110       End If
'
'120       EnumWindowsProc = True
'End Function
