VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MagicWand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dim undoch As changes
Dim parent As frmMain

Dim canmove As Boolean
Dim undoch As Changes

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    '   On Error GoTo MouseDown_Error

    curtilex = (parent.hScr.Value + X) \ parent.currenttilew
    curtiley = (parent.vScr.Value + Y) \ parent.currenttilew

    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
        Exit Sub
    End If

    Dim sourcetilenr As Integer
    If parent.sel.getIsInSelection(curtilex, curtiley) Then
        sourcetilenr = parent.sel.getSelTile(curtilex, curtiley)
    Else
        sourcetilenr = parent.getTile(curtilex, curtiley)
    End If

    If Button = vbRightButton Then
        Set undoch = New Changes
        parent.undoredo.ResetRedo

        Call parent.sel.ApplySelection(undoch, True)

        Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_APPLY)
        Exit Sub
    End If

    canmove = True

    If parent.sel.selstate = Append And (Button = vbLeftButton And parent.sel.getIsInSelection(curtilex, curtiley) And _
                                         Shift <> 1 And Shift <> 2) Then

        'start moving selection
        Call parent.sel.MouseDown(Button, Shift, X, Y)

    Else
        If Shift = 2 Then
            If parent.sel.getIsInSelection(curtilex, curtiley) Then
                parent.undoredo.ResetRedo
                Set undoch = New Changes

                Call Regions(False, sourcetilenr, curtilex, curtiley, frmGeneral.chkMagicWandScreen.value = vbChecked)
                Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_REMOVEAREA)
            End If
        ElseIf Shift = 1 Then
            If Not parent.sel.getIsInSelection(curtilex, curtiley) Then
                parent.undoredo.ResetRedo
                Set undoch = New Changes
            
                Call Regions(True, sourcetilenr, curtilex, curtiley, frmGeneral.chkMagicWandScreen.value = vbChecked)
                parent.sel.selstate = Append
            
                Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_ADDAREA)
            End If
        Else
            parent.undoredo.ResetRedo
            Set undoch = New Changes

            Call parent.sel.ApplySelection(undoch, False)
            Call Regions(True, sourcetilenr, curtilex, curtiley, frmGeneral.chkMagicWandScreen.value = vbChecked)
            parent.sel.selstate = Append
            Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_APPLY_AND_ADD)
        End If
        canmove = False
    End If

    '    If parent.sel.selstate = append Then
    '        If Button = vbRightButton Or _
             '           Button = vbLeftButton And parent.sel.getIsInSelection(curtilex, curtiley) And ((parent.pastetype <> p_normal And parent.sel.getSelTile(curtilex, curtiley) <> 0) Or parent.pastetype = p_normal) And _
             '           Shift <> 1 And Shift <> 2 Then
    '            Call parent.sel.MouseDown(Button, Shift, x, y)
    '            Exit Sub
    '        End If
    '    End If

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "MagicWand.MouseDown"
End Sub

Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Exit Sub

    Dim curtilex As Integer
    Dim curtiley As Integer
    curtilex = (parent.hScr.Value + X) \ parent.currenttilew
    curtiley = (parent.vScr.Value + Y) \ parent.currenttilew

    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
        Exit Sub
    End If

    If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) And _
       Shift <> 1 And Shift <> 2 Then
        parent.picPreview.MousePointer = 5
    Else
        parent.picPreview.MousePointer = 99
    End If

    If canmove And SharedVar.MouseDown <> 0 And parent.sel.selstate = Append Then
        Call parent.sel.MouseMove(Button, Shift, X, Y)
    End If
End Sub

Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If parent.sel.selstate = Append Then
        Call parent.sel.MouseUp(Button, Shift, X, Y)
    End If

End Sub

Sub Regions(add As Boolean, tilenr As Integer, ByVal X As Integer, ByVal Y As Integer, Optional fillscreenonly = True, Optional ByVal stack As Long = 50000)
    'Just like bucket fill, but adds (or remove) the targetted tiles to selection
    
    On Error GoTo Regions_Error

    Dim tick As Long
    tick = GetTickCount

    frmGeneral.IsBusy("MagicWand.Regions") = True

    Dim done(1023, 1023) As Boolean
    Dim px() As Long
    Dim py() As Long
    Dim Pstart As Long
    Dim Pstop As Long
    Dim Pcount As Long
    Dim Colour As Integer

    Dim lX As Integer
    Dim lY As Integer

    Dim Sloop As Long
    ReDim px(0)
    ReDim py(0)

    lX = X
    lY = Y


    Pstart = 0
    Pstop = 0
    px(0) = lX
    py(0) = lY
    Dim count As Long

    Dim lbsx As Integer
    Dim lbsy As Integer
    Dim hbsx As Integer
    Dim hbsy As Integer
    lbsx = parent.hScr.Value \ parent.currenttilew
    hbsx = (parent.hScr.Value + parent.picPreview.Width) \ parent.currenttilew
    lbsy = parent.vScr.Value \ parent.currenttilew
    hbsy = (parent.vScr.Value + parent.picPreview.Height) \ parent.currenttilew

    Call AddPoints(px, py, Pstart, Pstop, lX, lY, Sloop, stack)

    While Pstart <= Pstop And count < MAX_MAGICWAND

        lX = px(Pstart)
        lY = py(Pstart)
        If Not (lX < 0 Or lY < 0 Or lX > 1023 Or lY > 1023) Then
            If fillscreenonly Then
                If (lX < lbsx Or lY < lbsy Or lX > hbsx Or lY > hbsy) Then
                    GoTo skiploop
                End If
            End If

            If parent.sel.getIsInSelection(lX, lY) Then
                Colour = parent.sel.getSelTile(lX, lY)
            Else
                Colour = parent.getTile(lX, lY)
            End If

            'Consider special objects filling tiles (i.e. -22012...) as the same tile as object (i.e. 220)
            If Colour < 0 Then Colour = Colour \ -100

            If Not (done(lX, lY) = True Or Colour <> tilenr) Then
                Pcount = Pcount + 1
                done(lX, lY) = True

                count = count + 1

                If add Then
                    If Not parent.sel.getIsInSelection(lX, lY) Then
                        Call parent.sel.AddSelectionArea(lX, lX, lY, lY, undoch, False)
                    End If
                Else
                    If parent.sel.getIsInSelection(lX, lY) Then
                        Call parent.sel.RemoveSelectionArea(lX, lX, lY, lY, undoch, False)
                    End If
                End If

                Call AddPoints(px, py, Pstart, Pstop, lX, lY, Sloop, stack)
            End If
        End If
skiploop:
        Pstart = Pstart + 1

    Wend

    parent.AddDebug "MagicWand.Regions, " & count & " tiles number " & tilenr & " selected in " & GetTickCount - tick & " ms"
    If count >= MAX_MAGICWAND Then
        MessageBox "Selected area too large, operation aborted." & vbCrLf & "(" & count & " tiles number " & tilenr & " selected)", vbCritical + vbOKOnly, "Operation aborted"
    End If

    'That's unneeded, RemoveSelectionArea already updates selection state
'    If Not add Then
'        parent.sel.hasAlreadySelectedParts = parent.sel.areThereTilesInSelection
'    End If

    Call frmGeneral.UpdateToolBarButtons
    Call parent.UpdateLevel
    
    frmGeneral.IsBusy("MagicWand.Regions") = False

    On Error GoTo 0
    Exit Sub

Regions_Error:

    frmGeneral.IsBusy("MagicWand.Regions") = False
    
    HandleError Err, "MagicWand.Regions"
End Sub

Private Sub AddPoints(ByRef px() As Long, ByRef py() As Long, ByRef Pstart As Long, ByRef Pstop As Long, ByRef lX As Integer, ByRef lY As Integer, ByRef Sloop As Long, ByRef stack As Long)

    If Pstart > stack Then
        For Sloop = Pstart To Pstop
            px(Sloop - stack) = px(Sloop)
            py(Sloop - stack) = py(Sloop)
        Next Sloop
        ReDim Preserve px(Pstop - stack)
        ReDim Preserve py(Pstop - stack)
        Pstart = Pstart - stack
        Pstop = Pstop - stack
    End If
    If frmGeneral.chkMagicWandDiagonal.value = vbChecked Then
        ReDim Preserve px(Pstop + 8)
        ReDim Preserve py(Pstop + 8)
        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY

        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY + 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY + 1

        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY + 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY

        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY - 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY - 1

        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY - 1



    Else
        ReDim Preserve px(Pstop + 4)
        ReDim Preserve py(Pstop + 4)
        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY + 1
        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY - 1
    End If

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
