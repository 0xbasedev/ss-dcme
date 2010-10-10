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
10        AddDebug "--- Mbox prompt --- " & title & " - " & prompt & ";" & Buttons
20        result = MsgBox(prompt, Buttons, title)
30        AddDebug "--- Mbox result --- " & title & " - " & prompt & ";" & result
40        MessageBox = result

End Function



Sub AddDebug(str As String, Optional overrideLimit As Boolean = False)
      'Adds a line in the log file (dcme.log)
      'If the file exceeds the specified limit (MaxLogSize setting), the beginning is stripped
10        On Local Error Resume Next
          
        Dim maxsize As Long
        Dim f As Integer
        Dim path As String
        Dim temppath As String
        
20      path = App.path & "\DCME.log"
          
30      f = FreeFile

40      DebugInfo = DebugInfo & str & vbNewLine

50        maxsize = CLng(GetSetting("MaxLogSize", DEFAULT_MAX_LOG_SIZE))
          
60        If Not overrideLimit And settingsLoaded And FileExists(path) Then
70            If FileLen(path) > maxsize + LOG_BUFFER Then
                  'cut the first kb of the file
                  
80                f = FreeFile
                  
90                temppath = App.path & "\DCMEtmp" & f & ".log"
                  
100               Call RenameFile(path, temppath)
                  
110               Open temppath For Input As #f
                  
                  Dim longstring As String
                  
                  'seek to the 1024th byte and grab all data
120               Seek #f, LOG_BUFFER
130               longstring = Input(FileLen(temppath) - LOG_BUFFER, #f)

140               Close #f

                  'search the first newline in the grabbed text
150               longstring = Mid$(longstring, InStr(1, longstring, vbNewLine), Len(longstring))
                  
160               f = FreeFile
                  
                  'print that text
170               Open path For Output As #f
                  
180               Print #f, longstring
                  
190               Close #f
                  
200           End If
210       End If

          'print the new log line
220       Open path For Append As #f
          
          'newline is added automatically!
230       Print #f, str

240       Close #f

250       If temppath <> "" Then
260           DeleteFile (temppath)
270       End If

280       Debug.Print str
End Sub



Sub HandleError(ErrObj As ErrObject, Optional procedure As String = vbNullString, Optional showMsgBox As Boolean = True, Optional critical As Boolean = False)
          Const POSTYOURLOG As String = vbCrLf & "If the error persists, please post your log file (dcme.log) at http://www.ssforum.net in the DCME board."
          
10        frmGeneral.IsBusy(procedure) = False
          

              Dim errmsg As String
60            errmsg = ErrObj.Number & " (" & ErrObj.description & ") in " & procedure
              
70            If Erl <> 0 Then
80                errmsg = errmsg & " at line " & Erl
90            End If
100           If ErrObj.LastDllError <> 0 Then
110               errmsg = errmsg & " (Last DLL error: " & ErrObj.LastDllError & ")"
120           End If
              
130           If critical Then
140               Call AddDebug("*** Critical error " & errmsg, True)
150               MessageBox "Critical error " & errmsg & POSTYOURLOG, vbCritical + vbOKOnly
                  'attempt save
160               frmGeneral.Enabled = True
170               Call frmGeneral.QuickSaveAll
              
180               Call frmGeneral.CheckUnloadedForms(True)
          
190               UnHookWnd frmGeneral.hWnd
200               Unload frmGeneral
                  
210           Else
220               Call AddDebug("*** Error " & errmsg, True)
230               If showMsgBox Or bDEBUG Then MessageBox "Error " & errmsg & POSTYOURLOG, vbExclamation + vbOKOnly
240           End If

End Sub
