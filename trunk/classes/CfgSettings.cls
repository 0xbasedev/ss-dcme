VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CfgSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DEFAULTPATH = "\server.cfg"
Dim parent As frmMain
Dim cfgpath As String

Sub setParent(Main As frmMain)
    Set parent = Main
    If Not parent Is Nothing Then Call SearchCfg
End Sub

Sub SetCfgSetting(section As String, Key As String, newsetting As String)

End Sub

Function GetCfgSettingStr(section As String, Key As String) As String
    GetCfgSettingStr = INIload(section, Key, CStr(default), cfgpath)
End Function

Function GetCfgSetting(section As String, Key As String, Optional default As Long = 0) As Long
    GetCfgSetting = CLng(INIload(section, Key, CStr(default), cfgpath))
End Function


Function GetCfgPath() As String
    GetCfgPath = cfgpath
End Function

Sub SearchCfg()
    If FileExists(GetPathTo(parent.activeFile) & GetFileNameWithoutExtension(parent.activeFile) & ".cfg") Then
        cfgpath = GetPathTo(parent.activeFile) & GetFileNameWithoutExtension(parent.activeFile) & ".cfg"
    ElseIf FileExists(App.path & DEFAULTPATH) Then
        cfgpath = App.path & DEFAULTPATH
    Else
        cfgpath = ""
    End If
    
    UpdateSettingsLabel
End Sub

Sub SetCfgPath(path As String)
    If FileExists(path) Then
        cfgpath = path
    Else
        Call SearchCfg
    End If
    
    UpdateSettingsLabel
End Sub

Sub UpdateSettingsLabel()
    If cfgpath <> "" Then
        frmGeneral.lblCurrentSettings.Caption = "Settings: " & replace(cfgpath, App.path, "...", , 1)
    Else
        frmGeneral.lblCurrentSettings.Caption = "Settings: <DEFAULT>"
    End If
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    
End Sub
