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
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
          '   On Error GoTo MouseDown_Error

10        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
20        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

30        If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
40            Exit Sub
50        End If

          Dim sourcetilenr As Integer
60        If parent.sel.getIsInSelection(curtilex, curtiley) Then
70            sourcetilenr = parent.sel.getSelTile(curtilex, curtiley)
80        Else
90            sourcetilenr = parent.getTile(curtilex, curtiley)
100       End If

110       If Button = vbRightButton Then
120           Set undoch = New Changes
130           parent.undoredo.ResetRedo

140           Call parent.sel.ApplySelection(undoch, True)

150           Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_APPLY)
160           Exit Sub
170       End If

180       canmove = True

190       If parent.sel.selstate = Append And (Button = vbLeftButton And parent.sel.getIsInSelection(curtilex, curtiley) And _
                                               Shift <> 1 And Shift <> 2) Then

              'start moving selection
200           Call parent.sel.MouseDown(Button, Shift, X, Y)

210       Else
220           If Shift = 2 Then
230               If parent.sel.getIsInSelection(curtilex, curtiley) Then
240                   parent.undoredo.ResetRedo
250                   Set undoch = New Changes

260                   Call Regions(False, sourcetilenr, curtilex, curtiley, frmGeneral.chkMagicWandScreen.value = vbChecked)
270                   Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_REMOVEAREA)
280               End If
290           ElseIf Shift = 1 Then
300               If Not parent.sel.getIsInSelection(curtilex, curtiley) Then
310                   parent.undoredo.ResetRedo
320                   Set undoch = New Changes
                  
330                   Call Regions(True, sourcetilenr, curtilex, curtiley, frmGeneral.chkMagicWandScreen.value = vbChecked)
340                   parent.sel.selstate = Append
                  
350                   Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_ADDAREA)
360               End If
370           Else
380               parent.undoredo.ResetRedo
390               Set undoch = New Changes

400               Call parent.sel.ApplySelection(undoch, False)
410               Call Regions(True, sourcetilenr, curtilex, curtiley, frmGeneral.chkMagicWandScreen.value = vbChecked)
420               parent.sel.selstate = Append
430               Call parent.undoredo.AddToUndo(undoch, UNDO_WAND_APPLY_AND_ADD)
440           End If
450           canmove = False
460       End If

          '    If parent.sel.selstate = append Then
          '        If Button = vbRightButton Or _
                   '           Button = vbLeftButton And parent.sel.getIsInSelection(curtilex, curtiley) And ((parent.pastetype <> p_normal And parent.sel.getSelTile(curtilex, curtiley) <> 0) Or parent.pastetype = p_normal) And _
                   '           Shift <> 1 And Shift <> 2 Then
          '            Call parent.sel.MouseDown(Button, Shift, x, y)
          '            Exit Sub
          '        End If
          '    End If

470       On Error GoTo 0
480       Exit Sub

MouseDown_Error:
490       HandleError Err, "MagicWand.MouseDown"
End Sub

Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If Button = vbRightButton Then Exit Sub

          Dim curtilex As Integer
          Dim curtiley As Integer
20        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
30        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

40        If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
50            Exit Sub
60        End If

70        If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) And _
             Shift <> 1 And Shift <> 2 Then
80            parent.picPreview.MousePointer = 5
90        Else
100           parent.picPreview.MousePointer = 99
110       End If

120       If canmove And SharedVar.MouseDown <> 0 And parent.sel.selstate = Append Then
130           Call parent.sel.MouseMove(Button, Shift, X, Y)
140       End If
End Sub

Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If parent.sel.selstate = Append Then
20            Call parent.sel.MouseUp(Button, Shift, X, Y)
30        End If

End Sub

Sub Regions(add As Boolean, tilenr As Integer, ByVal X As Integer, ByVal Y As Integer, Optional fillscreenonly = True, Optional ByVal stack As Long = 50000)
          'Just like bucket fill, but adds (or remove) the targetted tiles to selection
          
10        On Error GoTo Regions_Error

          Dim tick As Long
20        tick = GetTickCount

30        frmGeneral.IsBusy("MagicWand.Regions") = True

          Dim done(1023, 1023) As Boolean
          Dim px() As Long
          Dim py() As Long
          Dim Pstart As Long
          Dim Pstop As Long
          Dim Pcount As Long
          Dim Colour As Integer

          Dim Lx As Integer
          Dim Ly As Integer

          Dim Sloop As Long
40        ReDim px(0)
50        ReDim py(0)

60        Lx = X
70        Ly = Y


80        Pstart = 0
90        Pstop = 0
100       px(0) = Lx
110       py(0) = Ly
          Dim count As Long

          Dim lbsx As Integer
          Dim lbsy As Integer
          Dim hbsx As Integer
          Dim hbsy As Integer
120       lbsx = parent.Hscr.value \ parent.currenttilew
130       hbsx = (parent.Hscr.value + parent.picPreview.width) \ parent.currenttilew
140       lbsy = parent.Vscr.value \ parent.currenttilew
150       hbsy = (parent.Vscr.value + parent.picPreview.height) \ parent.currenttilew

160       Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)

170       While Pstart <= Pstop And count < MAX_MAGICWAND

180           Lx = px(Pstart)
190           Ly = py(Pstart)
200           If Not (Lx < 0 Or Ly < 0 Or Lx > 1023 Or Ly > 1023) Then
210               If fillscreenonly Then
220                   If (Lx < lbsx Or Ly < lbsy Or Lx > hbsx Or Ly > hbsy) Then
230                       GoTo skiploop
240                   End If
250               End If

260               If parent.sel.getIsInSelection(Lx, Ly) Then
270                   Colour = parent.sel.getSelTile(Lx, Ly)
280               Else
290                   Colour = parent.getTile(Lx, Ly)
300               End If

                  'Consider special objects filling tiles (i.e. -22012...) as the same tile as object (i.e. 220)
310               If Colour < 0 Then Colour = Colour \ -100

320               If Not (done(Lx, Ly) = True Or Colour <> tilenr) Then
330                   Pcount = Pcount + 1
340                   done(Lx, Ly) = True

350                   count = count + 1

360                   If add Then
370                       If Not parent.sel.getIsInSelection(Lx, Ly) Then
380                           Call parent.sel.AddSelectionArea(Lx, Lx, Ly, Ly, undoch, False)
390                       End If
400                   Else
410                       If parent.sel.getIsInSelection(Lx, Ly) Then
420                           Call parent.sel.RemoveSelectionArea(Lx, Lx, Ly, Ly, undoch, False)
430                       End If
440                   End If

450                   Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)
460               End If
470           End If
skiploop:
480           Pstart = Pstart + 1

490       Wend

500       parent.AddDebug "MagicWand.Regions, " & count & " tiles number " & tilenr & " selected in " & GetTickCount - tick & " ms"
510       If count >= MAX_MAGICWAND Then
520           MessageBox "Selected area too large, operation aborted." & vbCrLf & "(" & count & " tiles number " & tilenr & " selected)", vbCritical + vbOKOnly, "Operation aborted"
530       End If

          'That's unneeded, RemoveSelectionArea already updates selection state
      '    If Not add Then
      '        parent.sel.hasAlreadySelectedParts = parent.sel.areThereTilesInSelection
      '    End If

540       Call frmGeneral.UpdateToolBarButtons
550       Call parent.UpdateLevel
          
560       frmGeneral.IsBusy("MagicWand.Regions") = False

570       On Error GoTo 0
580       Exit Sub

Regions_Error:

590       frmGeneral.IsBusy("MagicWand.Regions") = False
          
600       Call AddDebug("*** ERROR " & Err.Number & " (" & Err.description & ") in procedure Regions of Class Module MagicWand")
610       MessageBox "Error " & Err.Number & " (" & Err.description & ") in procedure Regions of Class Module MagicWand"

End Sub

Private Sub AddPoints(ByRef px() As Long, ByRef py() As Long, ByRef Pstart As Long, ByRef Pstop As Long, ByRef Lx As Integer, ByRef Ly As Integer, ByRef Sloop As Long, ByRef stack As Long)

10        If Pstart > stack Then
20            For Sloop = Pstart To Pstop
30                px(Sloop - stack) = px(Sloop)
40                py(Sloop - stack) = py(Sloop)
50            Next Sloop
60            ReDim Preserve px(Pstop - stack)
70            ReDim Preserve py(Pstop - stack)
80            Pstart = Pstart - stack
90            Pstop = Pstop - stack
100       End If
110       If frmGeneral.chkMagicWandDiagonal.value = vbChecked Then
120           ReDim Preserve px(Pstop + 8)
130           ReDim Preserve py(Pstop + 8)
140           Pstop = Pstop + 1
150           px(Pstop) = Lx + 1
160           py(Pstop) = Ly

170           Pstop = Pstop + 1
180           px(Pstop) = Lx + 1
190           py(Pstop) = Ly + 1
              '------
200           Pstop = Pstop + 1
210           px(Pstop) = Lx
220           py(Pstop) = Ly + 1

230           Pstop = Pstop + 1
240           px(Pstop) = Lx - 1
250           py(Pstop) = Ly + 1
              '------
260           Pstop = Pstop + 1
270           px(Pstop) = Lx - 1
280           py(Pstop) = Ly

290           Pstop = Pstop + 1
300           px(Pstop) = Lx - 1
310           py(Pstop) = Ly - 1
              '------
320           Pstop = Pstop + 1
330           px(Pstop) = Lx
340           py(Pstop) = Ly - 1

350           Pstop = Pstop + 1
360           px(Pstop) = Lx + 1
370           py(Pstop) = Ly - 1



380       Else
390           ReDim Preserve px(Pstop + 4)
400           ReDim Preserve py(Pstop + 4)
410           Pstop = Pstop + 1
420           px(Pstop) = Lx + 1
430           py(Pstop) = Ly
440           Pstop = Pstop + 1
450           px(Pstop) = Lx
460           py(Pstop) = Ly + 1
470           Pstop = Pstop + 1
480           px(Pstop) = Lx - 1
490           py(Pstop) = Ly
500           Pstop = Pstop + 1
510           px(Pstop) = Lx
520           py(Pstop) = Ly - 1
530       End If

          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
20        Set undoch = Nothing
End Sub
