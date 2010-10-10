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
10        Set parent = Main
20        If Not parent Is Nothing Then Call SearchCfg
End Sub

Sub SetCfgSetting(section As String, Key As String, newsetting As String)

End Sub

Function GetCfgSettingStr(section As String, Key As String) As String
10        GetCfgSettingStr = INIload(section, Key, CStr(default), cfgpath)
End Function

Function GetCfgSetting(section As String, Key As String, Optional default As Long = 0) As Long
10        GetCfgSetting = CLng(INIload(section, Key, CStr(default), cfgpath))
End Function


Function GetCfgPath() As String
10        GetCfgPath = cfgpath
End Function

Sub SearchCfg()
10        If FileExists(GetPathTo(parent.activeFile) & GetFileNameWithoutExtension(parent.activeFile) & ".cfg") Then
20            cfgpath = GetPathTo(parent.activeFile) & GetFileNameWithoutExtension(parent.activeFile) & ".cfg"
30        ElseIf FileExists(App.path & DEFAULTPATH) Then
40            cfgpath = App.path & DEFAULTPATH
50        Else
60            cfgpath = ""
70        End If
          
80        UpdateSettingsLabel
End Sub

Sub SetCfgPath(path As String)
10        If FileExists(path) Then
20            cfgpath = path
30        Else
40            Call SearchCfg
50        End If
          
60        UpdateSettingsLabel
End Sub

Sub UpdateSettingsLabel()
10        If cfgpath <> "" Then
20            frmGeneral.lblCurrentSettings.Caption = "Settings: " & replace(cfgpath, App.path, "...", , 1)
30        Else
40            frmGeneral.lblCurrentSettings.Caption = "Settings: <DEFAULT>"
50        End If
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
          
End Sub
