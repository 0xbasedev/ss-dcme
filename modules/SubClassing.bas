Attribute VB_Name = "SubClassing"
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

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

Private Const GWL_WNDPROC = -4
Private lpPrevWndProc As Long

Public otherinstance As Long

Public processingMessage As Boolean
Public startingGeneral As Boolean

Dim messageQueue() As Long
Dim messageQueueCount As Long

Global Const MSG_WINDOW_TITLE = "DCME_MESSAGE_HANDLER"

Private lpPrevGeneralWndProc As Long

Sub HookWnd(hWnd As Long)
    If Not bDEBUG Then
        lpPrevGeneralWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf GeneralWindowProc)
    End If
End Sub

Sub UnHookWnd(hWnd As Long)
    If Not bDEBUG Then
        If lpPrevGeneralWndProc Then
            SetWindowLong hWnd, GWL_WNDPROC, lpPrevGeneralWndProc
            lpPrevGeneralWndProc = 0
        End If
    End If
End Sub

Function GeneralWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim zDelta As Long
    Dim curpos As POINTAPI
    Select Case uMsg
    Case WM_DESTROY
        UnHookWnd hWnd

    Case WM_MOUSEWHEEL

        Call GetCursorPos(curpos)

        zDelta = (wParam \ 65535) Mod 65535


        If zDelta < 0 Then
            If frmGeneral.GetShift Then
                If curtool < T_lvz Then
                    frmGeneral.SetCurrentTool (curtool + 1)
                End If
            Else
                Call frmGeneral.ExecuteZoomFocus(True, curpos)
            End If
        Else
            If frmGeneral.GetShift Then
                If curtool > T_magnifier Then
                    frmGeneral.SetCurrentTool (curtool - 1)
                End If
            Else
                Call frmGeneral.ExecuteZoomFocus(False, curpos)
            End If
        End If
    Case Else
        GeneralWindowProc = CallWindowProc(lpPrevGeneralWndProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

Function HookMessageHandler() As Long
    If bDEBUG Then Exit Function
    
    otherinstance = CreateWindowEx(0, "STATIC", MSG_WINDOW_TITLE, 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)

    
    lpPrevWndProc = SetWindowLong(otherinstance, GWL_WNDPROC, AddressOf WindowProc)

    ReDim messageQueue(10)
    messageQueueCount = 0
        
    HookMessageHandler = otherinstance
End Function

Sub UnHookMessageHandler()
    If bDEBUG Then Exit Sub

    If lpPrevWndProc Then
        SetWindowLong otherinstance, GWL_WNDPROC, lpPrevWndProc
        lpPrevWndProc = 0
    End If

    If otherinstance Then
      Call DestroyWindow(otherinstance)
    End If
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
    Select Case uMsg
    Case WM_DESTROY
        UnHookMessageHandler

    Case WM_MOUSEWHEEL

        Call GetCursorPos(curpos)

        zDelta = (wParam \ 65535) Mod 65535


        If zDelta < 0 Then
            If frmGeneral.GetShift Then
                If curtool < T_lvz Then
                    frmGeneral.SetCurrentTool (curtool + 1)
                End If
            Else
                Call frmGeneral.ExecuteZoomFocus(True, curpos)
            End If
        Else
            If frmGeneral.GetShift Then
                If curtool > T_magnifier Then
                    frmGeneral.SetCurrentTool (curtool - 1)
                End If
            Else
                Call frmGeneral.ExecuteZoomFocus(False, curpos)
            End If
        End If

    Case WM_DCME_OPENMAP    'user-defined
        'wparam contains the pointer to the string we need
        'Dim path As String
        'path = PointerToString(wParam)
          AddDebug " +++ Received message '" & WM_DCME_OPENMAP & "', wParam=" & wParam & " lParam=" & lParam, True
          
          Call ProcessOpenMap(wParam, lParam)

        
        
        'Call PostMessage(lParam, &H10, 0, 0)
    Case Else
        WindowProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function


Private Sub AddQueue(wParam As Long)
    If UBound(messageQueue) >= messageQueueCount Then
        ReDim Preserve messageQueue(UBound(messageQueue) + 10)
    End If
    
    messageQueue(messageQueueCount) = wParam
    
    messageQueueCount = messageQueueCount + 1
End Sub


Sub ExecuteQueuedMessages()
    If messageQueueCount > 0 Then

        AddDebug "Processing queued message #1 / " & messageQueueCount & ": " & messageQueue(0)
        
        
        Dim toProcess As Long
        toProcess = messageQueue(0)
        
        'Remove it from queue
        Dim i As Long
        For i = 1 To messageQueueCount - 1
          messageQueue(i - 1) = messageQueue(i)
        Next
        messageQueueCount = messageQueueCount - 1
'
        Call ProcessOpenMap(toProcess, 0)
'
    End If
End Sub

Private Sub ProcessOpenMap(wParam As Long, lParam As Long)
        If processingMessage Or startingGeneral Then
          AddDebug "Message still being processed. Message queued #" & messageQueueCount & ": " & wParam
          Call AddQueue(wParam)
          Exit Sub
        End If
        
        On Error Resume Next
        
        


        Dim paths As String
        
'              If bDEBUG Then
'                paths = "C:\Jeux\Continuum\DCME\paths" & wParam
'              Else
        paths = App.path & "\paths" & wParam
'              End If
        
        
        
        If Not FileExists(paths) Then
            Exit Sub
        End If

        processingMessage = True

        
        Dim f As Integer
        f = FreeFile

        Open paths For Input As #f
        Dim str As String
        
        Line Input #f, str
        
        AddDebug " +++ Read '" & str & "' in '" & paths & "'", True
        
        Close #f
        
        AddDebug " +++ Deleting '" & paths & "'", True
        DeleteFile paths
        
      
'        On Local Error Resume Next
        
        'bring back DCME in front (this is also needed to avoid a crash caused when DCME is
        '                          minimized while a map opens)
'        screenUpdating = False
        
        Call frmGeneral.RestoreWindow
        frmGeneral.setfocus
        
        Dim mapidx As Integer
        
        mapidx = frmGeneral.GetIndexOfMap(str)
        
        If mapidx <> -1 Then
          AddDebug frmGeneral.hWnd & " +++ Map '" & str & "' already opened", True
          'map already opened
          Call frmGeneral.ActivateMap(mapidx)
        Else
          AddDebug frmGeneral.hWnd & " +++ Opening '" & str & "'", True
          Call frmGeneral.OpenMap(str)
          AddDebug frmGeneral.hWnd & " +++ OPENED '" & str & "'", True
  '        screenUpdating = True
        End If
        
        processingMessage = False
        
        Call ExecuteQueuedMessages
        
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
