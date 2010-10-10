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
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        On Error GoTo MouseDown_Error

20        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
30        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew
40        If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
50            Exit Sub
60        End If

70        If Button = vbLeftButton Or Button = vbRightButton Then
          
              Dim tmptile As Integer
              
80            If parent.sel.getIsInSelection(curtilex, curtiley) Then
90                tmptile = parent.sel.getSelTile(curtilex, curtiley)
100           Else
110               tmptile = parent.getTile(curtilex, curtiley)
120           End If
              
130           If tmptile < 0 Then
140               tmptile = tmptile \ -100
150           ElseIf tmptile = 0 Then
160               If frmGeneral.chkDropperIgnoreEmpty.value = vbChecked Then Exit Sub
170               tmptile = 256
180           End If
          
190           Call parent.tileset.SelectTiles(Button, tmptile, 1, 1, True)
              
200       End If

210       On Error GoTo 0
220       Exit Sub

MouseDown_Error:
230       HandleError Err, "Dropper.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
10        If SharedVar.MouseDown <> 0 Then
20            Call MouseDown(Button, X, Y)
30        End If
End Sub

Sub MouseUp()
    'TODO: Return to the previous tool after use
End Sub


Private Sub Class_Terminate()
10        Set parent = Nothing
End Sub
