Attribute VB_Name = "subMain"
Option Explicit

Dim ignoreMessageboxes As Boolean



Sub Main()
    bDEBUG = False
    Debug.Assert inDebug
    
    
    'what the fuck? Even in exe, it thinks it's in debug
'    bDEBUG = False
    
    otherinstance = FindWindow(vbNullString, MSG_WINDOW_TITLE)
   
   
'    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    
    
        
    ignoreMessageboxes = False
    

    If otherinstance <> 0 Then
        'Another instance is running
        Dim args As String
        
        Dim seq As Long
        Dim firsttick As Long
        firsttick = GetTickCount
        
        'This is to avoid having 2 files quickly made with the same name
        While GetTickCount = firsttick
            seq = GetTickCount
        Wend
        
        seq = GetTickCount
        
        args = command()
        
        If PrintPathsToLoad(args, seq) Then
            'Tell the other instance to load these files

            
            AddDebug "Sending DCME_OPENMAP (" & seq & ") to " & otherinstance
            'PostMessage otherinstance, WM_DCME_OPENMAP, seq, 0&
            SendMessageLong otherinstance, WM_DCME_OPENMAP, seq, 0
            
            AddDebug "SENT DCME_OPENMAP (" & seq & ") to " & otherinstance
            
        Else
            'There was nothing to load
        
        End If
        

        'Shut down this instance
        Exit Sub
        

    Else
        If CheckComponent("msinet.ocx") And _
            CheckComponent("comdlg32.ocx") And _
            CheckComponent("zlib.dll") And _
            CheckComponent("MSVBVM60.dll") And _
            CheckComponent("comct332.ocx") Then
               
               

            
            startingGeneral = True
            
            currentWindowsVersion = GetWindowsVersion
            
            Call HookMessageHandler
            'Messages can be received now
            
            Call LoadSettings
            
            'Check lvl file association
            If Not bDEBUG Then
                If Not IsLVLAssociatedToDCME Then
                    If GetSetting("AskForLVLAssociation", "Y") = "Y" Then
                        If MessageBox("Do you wish to associate Continuum map files (.LVL) with DCME? You can always do this manually from the preferences window if you don't want to now.", vbQuestion + vbYesNo) = vbYes Then
                            Call AssignExt
                        End If
                        Call SetSetting("AskForLVLAssociation", "N")
                    End If
                End If
            End If
            

            inSplash = True
'''
            frmSplash.show
            frmSplash.Refresh
'''
              
'''            frmSplash.TimerClose.Interval = 5000
            frmSplash.TimerClose.Enabled = True
'''
            
            If Not bDEBUG Then MakeTopMost frmSplash.hWnd
            
            Load frmGeneral
              
            If inSplash Then
                Unload frmSplash
            End If
            
            
'250               inSplash = False
    '        Unload Me
            startingGeneral = False
            
            Call ExecuteQueuedMessages
            
        Else
'290               inSplash = False
    '        Unload Me
        End If

    End If
End Sub


Private Function CheckComponent(filename As String) As Boolean
    If HaveComponent(filename) Then
        CheckComponent = True
    Else
        
        Dim result As VbMsgBoxResult
        
        result = MessageBox(filename & " was not found! Visit http://dcme.sscentral.com to get this file. Ignoring this error could cause unexpected crashes.", vbCritical + vbAbortRetryIgnore, "ERROR: Missing component")
        
        If result = vbAbort Then
            CheckComponent = False
            ignoreMessageboxes = True 'Don't pop up more msgboxes
        ElseIf result = vbIgnore Then
            'Ignore
            CheckComponent = True
        ElseIf result = vbRetry Then
            CheckComponent = CheckComponent(filename)
        End If
    End If
End Function


Private Function PrintPathsToLoad(args As String, seq As Long) As Boolean
    Dim paths As String
    paths = App.path & "\paths" & seq
    
    Dim fileext As String
    Dim longfilename As String
    Dim f As Integer
    f = FreeFile

    Open paths For Output As #f
    
    Dim tmp() As String
    
    tmp = Split(args, Chr(34))
    
    Dim i As Integer
    For i = 0 To UBound(tmp)

        
        longfilename = GetLongFilename(tmp(i))
        fileext = GetExtension(GetFileTitle(longfilename))
        
        If fileext = "lvl" Or fileext = "elvl" Or fileext = "bak" Then
            'if it's a lvl file, try loading it
            AddDebug "Printed '" & longfilename & "' to '" & paths & "'"
            PrintPathsToLoad = True
            Print #f, longfilename
        End If

    Next
    Close #f
End Function

'Private Function CheckOtherInstance() As Boolean
'
'    'Get window handles
'    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
'
'
'    If otherinstance <> 0 Then
'        'Another instance is running
'
'        'don't show ourselves if another instance is running
'        Dim args As String
'        args = Command()
'
'        Dim tmp() As String     'Splitted arguments
'        'split on quotes (")
'        tmp = Split(args, Chr(34))
'
'
'        Dim ret As Integer
'
'          Dim fileext As String
'          Dim longfilename As String
'        Dim f As Integer
'        f = FreeFile
'
'        Open App.path & "\paths" For Output As #f
'
'        For i = 0 To UBound(tmp)
'            'If i Mod 2 <> 0 Then
'
'            longfilename = GetLongFilename(tmp(i))
'            fileext = GetExtension(GetFileTitle(longfilename))
'
'            If GetExtension(GetFileTitle(GetLongFilename(tmp(i)))) = "lvl" Then
'                'if it's a lvl file, try loading it
'                Print #f, GetLongFilename(tmp(i))
'            End If
'            'End If
'        Next
'        Close #f
'
'        AddDebug "+++ openedMapByArgs " & openedmapbyargs, True
'
'        'already a DCME open and got args
'        If otherinstance <> 0 And openedmapbyargs Then
'            'Call SendMessageLong(otherinstance, WM_DCME_OPENMAP, StrPtr(str), Me.hwnd)
'
'            AddDebug Me.hWnd & " +++ Other instance found, sending OPENMAP message to " & otherinstance, True
'
'            Call SendMessageLong(otherinstance, WM_DCME_OPENMAP, 0&, 0&)
'
'            'UnHook Me.hWnd
'        End If
'        If otherinstance <> 0 Then
'            'Shutting down this instance
'            Call PostMessage(Me.hWnd, WM_CLOSE, 0, 0)
'        End If
'        DoEvents
'        Unload Me
'        End
'
'    Else
'
'
'End Function







'Function GetDebugLog() As String
'          'Max length of the text in a textbox
'          Const MAXTEXTLENGTH = 65535
'
'          Dim path As String
'          path = App.path & "\DCME.log"
'
'          'Returns the data contained in the log file
'10        If FileExists(path) Then
'
'              Dim f As Integer
'20            f = FreeFile
'30            Open path For Input As #f
'
'              Dim flen As Long
'
'              flen = FileLen(path)
'
'40            If flen > MAXTEXTLENGTH Then
'
'50                Seek #f, flen - MAXTEXTLENGTH
'60                GetDebugLog = Input(MAXTEXTLENGTH - 38, #f)
'                  'search the first newline in the grabbed text
'70                GetDebugLog = Mid$(GetDebugLog, InStr(1, GetDebugLog, vbNewLine), Len(GetDebugLog))
'
'80            Else
'90                GetDebugLog = Input(flen, #f)
'100           End If
'
'110           Close #f
'120       Else
'130           GetDebugLog = ""
'140       End If
'
'End Function
Sub ClearDebugLog()
      DeleteFile (App.path & "\DCME.log")
End Sub


Function Directory_Temp() As String
    Directory_Temp = App.path & "\DCME temporary files"
End Function

Function Directory_Cache() As String
    Directory_Cache = App.path & "\DCME temporary files\cache"
End Function
