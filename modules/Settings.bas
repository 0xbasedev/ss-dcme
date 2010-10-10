Attribute VB_Name = "Settings"
Option Explicit

Dim nrOfSettings As Integer
Dim keys() As String
Dim settings() As String

Public settingsLoaded As Boolean

Sub SaveSettings()
    If Not settingsLoaded Then
        'make sure settings have been loaded (so we don't save settings before we got any data)
        Exit Sub
    End If
    
    Dim f As Integer
    f = FreeFile
    
    If DeleteFile(App.path & "\settings.dat") Then
    
    
'    If FileExists(App.path & "\settings.dat") Then
'        Kill App.path & "\settings.dat"
'    End If

        Open App.path & "\settings.dat" For Binary As #f
        Put #f, , nrOfSettings
        Put #f, , keys
        Put #f, , settings
        Close #f
    Else
        If bDEBUG Then
        
        End If
    End If
End Sub

Sub LoadSettings()
    If Not FileExists(App.path & "\settings.dat") Then
        nrOfSettings = 0
        ReDim keys(0)
        ReDim settings(0)
        settingsLoaded = True
        Exit Sub
    End If

    Dim f As Integer
    f = FreeFile
    Open App.path & "\settings.dat" For Binary As #f
    Get #f, , nrOfSettings
    ReDim keys(nrOfSettings)
    ReDim settings(nrOfSettings)
    Get #f, , keys
    Get #f, , settings
    Close #f
    
    settingsLoaded = True
End Sub

Sub ClearSettings()
    Call DeleteFile(App.path & "\settings.dat")
    Call LoadSettings
End Sub

Function GetSetting(Key As String, Optional defaultval As Variant = vbNullString) As String
    If Not settingsLoaded Then
        GetSetting = defaultval
        Exit Function
    End If
    
    Dim Index As Integer
    Index = getSettingIndex(Key)

    If Index = -1 Then
        GetSetting = defaultval
    Else
        GetSetting = settings(Index)
    End If
End Function

Sub SetSetting(Key As String, value As String)
    Dim Index As Integer
    Index = getSettingIndex(Key)

    If Index = -1 Then
        Call CreateSetting(Key, value)
    Else
        settings(Index) = value
    End If
End Sub

Private Sub CreateSetting(Key As String, value As String)
    ReDim Preserve settings(UBound(settings) + 1)
    ReDim Preserve keys(UBound(keys) + 1)
    settings(UBound(settings)) = value
    keys(UBound(keys)) = Key & ":" & UBound(settings)
    nrOfSettings = nrOfSettings + 1
End Sub

Sub RemoveSetting(Key As String)
    Dim Index As Integer
    Index = getSettingIndex(Key)

    If Index = -1 Then
        'no entry to be removed
    Else
        Call RemoveKey(Key)

        Dim j As Integer
        'move all other settings
        For j = Index To UBound(settings) - 1
            settings(j) = settings(j + 1)
        Next
        'remove the last setting,empty setting, by shrinking the array
        ReDim Preserve settings(UBound(settings) - 1)
        nrOfSettings = nrOfSettings - 1
    End If
End Sub

Private Sub RemoveKey(Key As String)
    Dim i As Integer
    Dim Index As Integer
    Index = -1
    For i = 0 To UBound(keys)
        Dim parts() As String
        parts = Split(keys(i), ":")
        If LCase(parts(0)) = LCase(Key) Then
            Index = i
            Exit For
        End If
    Next

    If Index = -1 Then
        'no key found
    Else
        Dim j As Integer
        'move all other keys
        For j = Index To UBound(keys) - 1
            keys(j) = keys(j + 1)
        Next
        'remove the last key by shrinking the array
        ReDim Preserve keys(UBound(keys) - 1)
    End If
End Sub

Private Function getSettingIndex(Key As String)
    Dim i As Integer
    For i = 0 To UBound(keys)
        If keys(i) <> "" Then
            Dim parts() As String
            parts = Split(keys(i), ":")
            If parts(0) = Key Then
                getSettingIndex = CInt(parts(1))
                Exit Function
            End If
        End If
    Next

    'the for didn't find the key, so it doesn't exist
    getSettingIndex = -1
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
