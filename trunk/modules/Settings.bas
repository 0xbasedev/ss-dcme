Attribute VB_Name = "Settings"
Option Explicit

Dim nrOfSettings As Integer
Dim keys() As String
Dim settings() As String

Public settingsLoaded As Boolean

Sub SaveSettings()
10        If Not settingsLoaded Then
              'make sure settings have been loaded (so we don't save settings before we got any data)
20            Exit Sub
30        End If
          
          Dim f As Integer
40        f = FreeFile
          
50        If DeleteFile(App.path & "\settings.dat") Then
          
          
      '    If FileExists(App.path & "\settings.dat") Then
      '        Kill App.path & "\settings.dat"
      '    End If

60            Open App.path & "\settings.dat" For Binary As #f
70            Put #f, , nrOfSettings
80            Put #f, , keys
90            Put #f, , settings
100           Close #f
110       Else
120           If bDEBUG Then
              
130           End If
140       End If
End Sub

Sub LoadSettings()
10        If Not FileExists(App.path & "\settings.dat") Then
20            nrOfSettings = 0
30            ReDim keys(0)
40            ReDim settings(0)
50            settingsLoaded = True
60            Exit Sub
70        End If

          Dim f As Integer
80        f = FreeFile
90        Open App.path & "\settings.dat" For Binary As #f
100       Get #f, , nrOfSettings
110       ReDim keys(nrOfSettings)
120       ReDim settings(nrOfSettings)
130       Get #f, , keys
140       Get #f, , settings
150       Close #f
          
160       settingsLoaded = True
End Sub



Function GetSetting(Key As String, Optional defaultval As Variant = vbNullString) As String
10        If Not settingsLoaded Then
20            GetSetting = defaultval
30            Exit Function
40        End If
          
          Dim Index As Integer
50        Index = getSettingIndex(Key)

60        If Index = -1 Then
70            GetSetting = defaultval
80        Else
90            GetSetting = settings(Index)
100       End If
End Function

Sub SetSetting(Key As String, value As String)
          Dim Index As Integer
10        Index = getSettingIndex(Key)

20        If Index = -1 Then
30            Call CreateSetting(Key, value)
40        Else
50            settings(Index) = value
60        End If
End Sub

Private Sub CreateSetting(Key As String, value As String)
10        ReDim Preserve settings(UBound(settings) + 1)
20        ReDim Preserve keys(UBound(keys) + 1)
30        settings(UBound(settings)) = value
40        keys(UBound(keys)) = Key & ":" & UBound(settings)
50        nrOfSettings = nrOfSettings + 1
End Sub

Sub RemoveSetting(Key As String)
          Dim Index As Integer
10        Index = getSettingIndex(Key)

20        If Index = -1 Then
              'no entry to be removed
30        Else
40            Call RemoveKey(Key)

              Dim j As Integer
              'move all other settings
50            For j = Index To UBound(settings) - 1
60                settings(j) = settings(j + 1)
70            Next
              'remove the last setting,empty setting, by shrinking the array
80            ReDim Preserve settings(UBound(settings) - 1)
90            nrOfSettings = nrOfSettings - 1
100       End If
End Sub

Private Sub RemoveKey(Key As String)
          Dim i As Integer
          Dim Index As Integer
10        Index = -1
20        For i = 0 To UBound(keys)
              Dim parts() As String
30            parts = Split(keys(i), ":")
40            If LCase(parts(0)) = LCase(Key) Then
50                Index = i
60                Exit For
70            End If
80        Next

90        If Index = -1 Then
              'no key found
100       Else
              Dim j As Integer
              'move all other keys
110           For j = Index To UBound(keys) - 1
120               keys(j) = keys(j + 1)
130           Next
              'remove the last key by shrinking the array
140           ReDim Preserve keys(UBound(keys) - 1)
150       End If
End Sub

Private Function getSettingIndex(Key As String)
          Dim i As Integer
10        For i = 0 To UBound(keys)
20            If keys(i) <> "" Then
                  Dim parts() As String
30                parts = Split(keys(i), ":")
40                If parts(0) = Key Then
50                    getSettingIndex = CInt(parts(1))
60                    Exit Function
70                End If
80            End If
90        Next

          'the for didn't find the key, so it doesn't exist
100       getSettingIndex = -1
End Function





Function GetLastDialogPath(dialog As String) As String
    Dim ret As String
    
    ret = GetSetting("dlg" & dialog)
    
    If ret = "" Then
        GetLastDialogPath = App.path
        
    ElseIf DirExists(ret) Then
        GetLastDialogPath = ret
        
    Else
        GetLastDialogPath = ""
    End If
End Function

Sub SetLastDialogPath(dialog As String, path As String)
    Call SetSetting("dlg" & dialog, path)
End Sub
