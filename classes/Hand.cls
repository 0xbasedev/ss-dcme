VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim firstDragX As Integer, firstDragY As Integer

Dim firstValX As Integer, firstValY As Integer


Dim parent As frmMain

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(X As Single, Y As Single)
'    dragxcorrection = X \ parent.currenttilew
'    dragycorrection = Y \ parent.currenttilew
'    oldDragX = dragxcorrection
'    oldDragY = dragycorrection
    
    firstDragX = X
    firstDragY = Y
    
    firstValX = parent.hScr.value
    firstValY = parent.vScr.value
    
    
End Sub

Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button Then

        Call parent.SetScrollbarValues(firstValX - (X - firstDragX), firstValY - (Y - firstDragY), False)
        
        Call parent.UpdateLevel(True, True)
    End If

End Sub

Sub MouseUp()
'update level properly (in case something weird happened in the fast update
    Call parent.UpdateLevel(False)

End Sub


Private Sub Class_Terminate()
    Set parent = Nothing
End Sub
