Attribute VB_Name = "mdlDebug"
Option Explicit

Dim DebugInfo As String

Global Const DEFAULT_MAX_LOG_SIZE = "1048576" ' = ~1MB
'number of bytes that can exceed
'the set limit of log file size before
'it is cut down
Global Const LOG_BUFFER = 1024




Function MessageBox(prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional title As String = "DCME") As VbMsgBoxResult
  
  If inSplash Then Unload frmSplash
  
    Dim result As VbMsgBoxResult
    AddDebug "--- Mbox prompt --- " & title & " - " & prompt & ";" & Buttons
    result = MsgBox(prompt, Buttons, title)
    AddDebug "--- Mbox result --- " & title & " - " & prompt & ";" & result
    MessageBox = result

End Function



Sub AddDebug(str As String, Optional overrideLimit As Boolean = False)
'Adds a line in the log file (dcme.log)
'If the file exceeds the specified limit (MaxLogSize setting), the beginning is stripped
    On Local Error Resume Next
    
  Dim maxsize As Long
  Dim f As Integer
  Dim path As String
  Dim temppath As String
  
  path = App.path & "\DCME.log"
    
  f = FreeFile

  DebugInfo = DebugInfo & str & vbNewLine

    maxsize = CLng(GetSetting("MaxLogSize", DEFAULT_MAX_LOG_SIZE))
    
    If Not overrideLimit And settingsLoaded And FileExists(path) Then
        If FileLen(path) > maxsize + LOG_BUFFER Then
            'cut the first kb of the file
            
            f = FreeFile
            
            temppath = App.path & "\DCMEtmp" & f & ".log"
            
            Call RenameFile(path, temppath)
            
            Open temppath For Input As #f
            
            Dim longstring As String
            
            'seek to the 1024th byte and grab all data
            Seek #f, LOG_BUFFER
            longstring = Input(FileLen(temppath) - LOG_BUFFER, #f)

            Close #f

            'search the first newline in the grabbed text
            longstring = Mid$(longstring, InStr(1, longstring, vbNewLine), Len(longstring))
            
            f = FreeFile
            
            'print that text
            Open path For Output As #f
            
            Print #f, longstring
            
            Close #f
            
        End If
    End If

    'print the new log line
    Open path For Append As #f
    
    'newline is added automatically!
    Print #f, str

    Close #f

    If temppath <> "" Then
        DeleteFile (temppath)
    End If

    Debug.Print str
End Sub



Sub HandleError(ErrObj As ErrObject, Optional procedure As String = vbNullString, Optional showMsgBox As Boolean = True, Optional critical As Boolean = False)
    Const POSTYOURLOG As String = vbCrLf & "If the error persists, please post your log file (dcme.log) at http://www.ssforum.net in the DCME board."
    
    frmGeneral.IsBusy(procedure) = False
    

        Dim errmsg As String
        errmsg = ErrObj.Number & " (" & ErrObj.description & ") in " & procedure
        
        If Erl <> 0 Then
            errmsg = errmsg & " at line " & Erl
        End If
        If ErrObj.LastDllError <> 0 Then
            errmsg = errmsg & " (Last DLL error: " & ErrObj.LastDllError & ")"
        End If
        
        If critical Then
            Call AddDebug("*** Critical error " & errmsg, True)
            MessageBox "Critical error " & errmsg & POSTYOURLOG, vbCritical + vbOKOnly
            'attempt save
            frmGeneral.Enabled = True
            Call frmGeneral.QuickSaveAll
        
            Call frmGeneral.CheckUnloadedForms(True)
    
            UnHookWnd frmGeneral.hWnd
            Unload frmGeneral
            
        Else
            Call AddDebug("*** Error " & errmsg, True)
            If showMsgBox Or bDEBUG Then MessageBox "Error " & errmsg & POSTYOURLOG, vbExclamation + vbOKOnly
        End If

End Sub
