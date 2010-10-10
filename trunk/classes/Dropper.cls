VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dropper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    On Error GoTo MouseDown_Error

    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew
    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
        Exit Sub
    End If

    If Button = vbLeftButton Or Button = vbRightButton Then
    
        Dim tmptile As Integer
        
        If parent.sel.getIsInSelection(curtilex, curtiley) Then
            tmptile = parent.sel.getSelTile(curtilex, curtiley)
        Else
            tmptile = parent.getTile(curtilex, curtiley)
        End If
        
        If tmptile < 0 Then
            tmptile = tmptile \ -100
        ElseIf tmptile = 0 Then
            If frmGeneral.chkDropperIgnoreEmpty.value = vbChecked Then Exit Sub
            tmptile = 256
        End If
    
        Call parent.tileset.SelectTiles(Button, tmptile, 1, 1, True)
        
    End If

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "Dropper.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
    If SharedVar.MouseDown <> 0 Then
        Call MouseDown(Button, X, Y)
    End If
End Sub

Sub MouseUp()
    'TODO: Return to the previous tool after use
End Sub


Private Sub Class_Terminate()
    Set parent = Nothing
End Sub
