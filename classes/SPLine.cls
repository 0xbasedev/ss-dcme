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
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    On Error GoTo MouseDown_Error

    Dim step As Integer
    step = frmGeneral.toolStep(curtool - 1).value

    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023
    
    
    If Button = vbLeftButton Then
        
        If SharedVar.splineInProgress Then
            'Not the first click, draw a new line
            
            Call parent.tline.DrawLine(startlinex, startliney, curtilex, curtiley, undoch, False, True, True)
            Call parent.UpdateLevel(False, False)
        Else
            'First left-click, get tileset info
            Dim toolOptions As DrawOptions
            Dim curSelection As TilesetSelection
            
            curSelection = parent.tileset.selection(Button)
            
            ''''''''''''''''''
            'get tool options'
            If curtool = T_customshape Then
                toolOptions.step = 0
            Else
                toolOptions.step = frmGeneral.toolStep(curtool - 1).value
            End If
            
            If curSelection.tileSize.X <> 1 Or curSelection.tileSize.Y <> 1 Then
                toolOptions.size = 1
            Else
                toolOptions.size = frmGeneral.toolSize(curtool - 1).value
            End If
            
            If frmGeneral.optToolRound(SharedVar.curtool - 1).value Then
                toolOptions.drawshape = DS_Circle
            Else
                toolOptions.drawshape = DS_Rectangle
            End If
            
            'set tileset selection
            Call parent.tline.SetSelection(curSelection)
            'set options
            Call parent.tline.SetOptions(toolOptions)
            
            
            'Init undo stack
            parent.undoredo.ResetRedo
            Set undoch = New Changes
            
            
            'It has begun
            SharedVar.splineInProgress = True
            
        End If
        
    ElseIf Button = vbRightButton Then
        'Right-click, end line
        If SharedVar.splineInProgress Then
            Call parent.undoredo.AddToUndo(undoch, UNDO_SPLINE)
            SharedVar.splineInProgress = False
        End If
    End If


    startlinex = curtilex
    startliney = curtiley


    If parent.magnifier.UsingPixels Then
        parent.UpdateLevel
    Else
        parent.UpdatePreview
    End If

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "SPLine.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    On Error GoTo MouseMove_Error


    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023

    If endlinex = curtilex And endliney = curtiley Then
        Exit Sub
    End If
    
    'Draw preview
    If SharedVar.splineInProgress Then
        'Draw preview of a line
        Call parent.tline.DrawLine(startlinex, startliney, curtilex, curtiley, undoch, True, True, False)
    End If

    Call parent.UpdatePreview(False)

    endlinex = curtilex
    endliney = curtiley

    On Error GoTo 0
    Exit Sub

MouseMove_Error:
    HandleError Err, "SPLine.MouseMove"
End Sub

Sub MouseUp(Button As Integer, X As Single, Y As Single)
    'nothing to do here
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
