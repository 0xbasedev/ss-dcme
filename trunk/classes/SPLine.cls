VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Dim startlinex As Integer
Dim startliney As Integer
Dim endlinex As Integer
Dim endliney As Integer

Dim undoch As Changes

Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        On Error GoTo MouseDown_Error

          Dim step As Integer
20        step = frmGeneral.toolStep(curtool - 1).value

30        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
40        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

50        If curtilex < 0 Then curtilex = 0
60        If curtiley < 0 Then curtiley = 0
70        If curtilex > 1023 Then curtilex = 1023
80        If curtiley > 1023 Then curtiley = 1023
          
          
90        If Button = vbLeftButton Then
              
100           If SharedVar.splineInProgress Then
                  'Not the first click, draw a new line
                  
110               Call parent.tline.DrawLine(startlinex, startliney, curtilex, curtiley, undoch, False, True, True)
120               Call parent.UpdateLevel(False, False)
130           Else
                  'First left-click, get tileset info
                  Dim toolOptions As DrawOptions
                  Dim curSelection As TilesetSelection
                  
140               curSelection = parent.tileset.selection(Button)
                  
                  ''''''''''''''''''
                  'get tool options'
150               If curtool = T_customshape Then
160                   toolOptions.step = 0
170               Else
180                   toolOptions.step = frmGeneral.toolStep(curtool - 1).value
190               End If
                  
200               If curSelection.tileSize.X <> 1 Or curSelection.tileSize.Y <> 1 Then
210                   toolOptions.size = 1
220               Else
230                   toolOptions.size = frmGeneral.toolSize(curtool - 1).value
240               End If
                  
250               If frmGeneral.optToolRound(SharedVar.curtool - 1).value Then
260                   toolOptions.drawshape = DS_Circle
270               Else
280                   toolOptions.drawshape = DS_Rectangle
290               End If
                  
                  'set tileset selection
300               Call parent.tline.SetSelection(curSelection)
                  'set options
310               Call parent.tline.SetOptions(toolOptions)
                  
                  
                  'Init undo stack
320               parent.undoredo.ResetRedo
330               Set undoch = New Changes
                  
                  
                  'It has begun
340               SharedVar.splineInProgress = True
                  
350           End If
              
360       ElseIf Button = vbRightButton Then
              'Right-click, end line
370           If SharedVar.splineInProgress Then
380               Call parent.undoredo.AddToUndo(undoch, UNDO_SPLINE)
390               SharedVar.splineInProgress = False
400           End If
410       End If


420       startlinex = curtilex
430       startliney = curtiley


440       If parent.magnifier.UsingPixels Then
450           parent.UpdateLevel
460       Else
470           parent.UpdatePreview
480       End If

490       On Error GoTo 0
500       Exit Sub

MouseDown_Error:
510       HandleError Err, "SPLine.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        On Error GoTo MouseMove_Error


20        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
30        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

40        If curtilex < 0 Then curtilex = 0
50        If curtiley < 0 Then curtiley = 0
60        If curtilex > 1023 Then curtilex = 1023
70        If curtiley > 1023 Then curtiley = 1023

80        If endlinex = curtilex And endliney = curtiley Then
90            Exit Sub
100       End If
          
          'Draw preview
110       If SharedVar.splineInProgress Then
              'Draw preview of a line
120           Call parent.tline.DrawLine(startlinex, startliney, curtilex, curtiley, undoch, True, True, False)
130       End If

140       Call parent.UpdatePreview(False)

150       endlinex = curtilex
160       endliney = curtiley

170       On Error GoTo 0
180       Exit Sub

MouseMove_Error:
190       HandleError Err, "SPLine.MouseMove"
End Sub

Sub MouseUp(Button As Integer, X As Single, Y As Single)
    'nothing to do here
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
20        Set undoch = Nothing
End Sub
