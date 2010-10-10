VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreehandSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain
Dim c_isdrawing As Boolean
Dim undoch As Changes

Private Type Coordinate
    X As Integer
    Y As Integer
End Type

Dim pts() As Coordinate
Dim PtCount As Long
Dim Boundaries As area

Dim previousX As Integer
Dim previousY As Integer

Private Enum checkedType
    setTrue = -1
    notChecked = 0
    setFalse = 1
    isedge = 2
    checkedEdge = 3
End Enum

Dim bitfield() As checkedType

Dim movingselection As Boolean

Dim dragxcorrection As Integer
Dim dragycorrection As Integer

Public Property Get isDrawing() As Boolean
10        isDrawing = c_isdrawing
End Property

Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        curtilex = (parent.hScr.value + X) \ parent.currenttilew
20        curtiley = (parent.vScr.value + Y) \ parent.currenttilew
30        If curtilex < 0 Then curtilex = 0
40        If curtiley < 0 Then curtiley = 0
50        If curtilex > 1023 Then curtilex = 1023
60        If curtiley > 1023 Then curtiley = 1023

70        If parent.sel.hasAlreadySelectedParts Then
              'check if we click inside the selection or outside, and with which button
80            If Button = vbRightButton Then
90                Set undoch = New Changes
100               parent.undoredo.ResetRedo

110               Call parent.sel.ApplySelection(undoch, True)

120               Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_APPLY)

130           ElseIf Button = vbLeftButton Then
                  'If isInSelection(curtilex, curtiley) And _
                   '    Shift <> 1 And Shift <> 2 Then
140               If parent.sel.getIsInSelection(curtilex, curtiley) And ((parent.pastetype <> p_normal And parent.sel.getSelTile(curtilex, curtiley) <> 0) Or parent.pastetype = p_normal) And _
                     Shift <> 1 And Shift <> 2 Then
                      'we clicked in the selection, ready to move
150                   dragxcorrection = curtilex
160                   dragycorrection = curtiley
170                   previousX = curtilex
180                   previousY = curtiley
190                   movingselection = True
                    
                        Call parent.sel.StartMove
200               Else
210                   Set undoch = New Changes
220                   parent.undoredo.ResetRedo
                      
230                   If Shift <> 1 And Shift <> 2 Then
240                       Call parent.sel.ApplySelection(undoch, True)
250                   End If

                      'we want to draw if we clicked outside of selection or
                      ' if we use shift or ctrl

                      'we want to draw again
260                   Call ClearPoints
270                   Call AddPoint(curtilex, curtiley, pts, PtCount)
                      
280                   Boundaries.Right = curtilex
290                   Boundaries.Left = curtilex
300                   Boundaries.Bottom = curtiley
310                   Boundaries.Top = curtiley
                              
320                   previousX = curtilex
330                   previousY = curtiley
340                   c_isdrawing = True
                      
350                   Call parent.UpdatePreview(True, False)

360               End If
370           End If
380       Else
390           Set undoch = New Changes
400           parent.undoredo.ResetRedo

              'we want to draw if we clicked outside of selection or
              ' if we use shift or ctrl

              'we want to draw again
410           Call ClearPoints
420           Call AddPoint(curtilex, curtiley, pts, PtCount)
              
430           Boundaries.Right = curtilex
440           Boundaries.Left = curtilex
450           Boundaries.Bottom = curtiley
460           Boundaries.Top = curtiley
                      
470           previousX = curtilex
480           previousY = curtiley
490           c_isdrawing = True
              
500           Call parent.UpdatePreview(True, False)
510       End If

End Sub

Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
          
10        curtilex = (parent.hScr.value + X) \ parent.currenttilew
20        curtiley = (parent.vScr.value + Y) \ parent.currenttilew
30        If curtilex < 0 Then curtilex = 0
40        If curtiley < 0 Then curtiley = 0
50        If curtilex > 1023 Then curtilex = 1023
60        If curtiley > 1023 Then curtiley = 1023

70        If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) And ((parent.pastetype <> p_normal And parent.sel.getSelTile(curtilex, curtiley) <> 0) Or parent.pastetype = p_normal) And _
             Shift <> 1 And Shift <> 2 Then
80            parent.picPreview.MousePointer = 5
90        Else
100           parent.picPreview.MousePointer = 2
110       End If
          
120       If c_isdrawing Then
130           If curtilex <> previousX Or curtiley <> previousY Then
                  
140               Call AddPoint(curtilex, curtiley, pts, PtCount)

150               If curtilex > Boundaries.Right Then Boundaries.Right = curtilex
160               If curtilex < Boundaries.Left Then Boundaries.Left = curtilex
170               If curtiley > Boundaries.Bottom Then Boundaries.Bottom = curtiley
180               If curtiley < Boundaries.Top Then Boundaries.Top = curtiley
                  
190               RedrawPoints True
                  
200               previousX = curtilex
210               previousY = curtiley
220           End If
              
              
230       ElseIf movingselection Then
240           If Button = vbLeftButton Then
                  'we are dragging the selection around
250               Call parent.sel.Move(curtilex - dragxcorrection, curtiley - dragycorrection)

260               dragxcorrection = curtilex
270               dragycorrection = curtiley
280           End If
290       End If
          
          
End Sub



Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
10        frmGeneral.IsBusy("FreehandSelection.MouseUp") = True
          
20        If c_isdrawing Then
30            Set undoch = New Changes
40            Call parent.undoredo.ResetRedo
              
50            If PtCount > 1 Then
60                Call AddPoint(pts(0).X, pts(0).Y, pts, PtCount)
70            End If
80            RedrawPoints True
              
90            Call RegionFromPoints(Not IsControl(Shift))
          
100           c_isdrawing = False
              
110           Call parent.undoredo.AddToUndo(undoch, "Freehand Selection")
              
120           Call parent.UpdateLevel
130       ElseIf movingselection Then
          
              Dim dx As Integer
              Dim dy As Integer
              Dim curtilex As Integer
              Dim curtiley As Integer
              
              'calculate tile coordinates
140           curtilex = (parent.hScr.value + X) \ parent.currenttilew
150           curtiley = (parent.vScr.value + Y) \ parent.currenttilew

160           movingselection = False

170           dx = curtilex - previousX
180           dy = curtiley - previousY
190           If dx <> 0 And dy <> 0 Then
200               parent.undoredo.ResetRedo
210               Set undoch = New Changes

                    Call parent.sel.EndMove(undoch)

230               Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_MOVE)
240           End If
              
250       End If
          
260       frmGeneral.IsBusy("FreehandSelection.MouseUp") = False
          
End Sub

Sub RedrawPoints(Refresh As Boolean)
          Dim i As Integer
          Dim maplbx As Integer
          Dim maplby As Integer
          Dim maphbx As Integer
          Dim maphby As Integer

10        maplbx = parent.hScr.value \ parent.currenttilew
20        maplby = parent.vScr.value \ parent.currenttilew
30        maphbx = maplbx + parent.picPreview.width \ parent.currenttilew
40        maphby = maplby + parent.picPreview.height \ parent.currenttilew
          
50        For i = 1 To PtCount - 1
            Call DrawLine(parent.hBufferDC, pts(i).X * parent.currenttilew - parent.hScr.value + parent.currenttilew \ 2, _
                                    pts(i).Y * parent.currenttilew - parent.vScr.value + parent.currenttilew \ 2, _
                                    pts(i - 1).X * parent.currenttilew - parent.hScr.value + parent.currenttilew \ 2, _
                                    pts(i - 1).Y * parent.currenttilew - parent.vScr.value + parent.currenttilew \ 2, _
                                    vbWhite)



'60            parent.picPreview.Line (pts(i).X * parent.currenttilew - parent.hScr.value + parent.currenttilew \ 2, pts(i).Y * parent.currenttilew - parent.vScr.value + parent.currenttilew \ 2)-(pts(i - 1).X * parent.currenttilew - parent.hScr.value + parent.currenttilew \ 2, pts(i - 1).Y * parent.currenttilew - parent.vScr.value + parent.currenttilew \ 2)
70        Next

80        If Refresh Then parent.UpdatePreview
End Sub

Private Sub AddPoint(X As Integer, Y As Integer, ByRef pts() As Coordinate, ByRef count As Long)
10        If count > UBound(pts) Then
20            ReDim Preserve pts(UBound(pts) + 500)
30        End If
          
40        pts(count).X = X
50        pts(count).Y = Y
          
60        count = count + 1
End Sub

Private Sub ClearPoints()
10        PtCount = 0
20        Erase pts
30        ReDim pts(0)
40        Boundaries.Right = -1
50        Boundaries.Left = 1024
60        Boundaries.Top = 1024
70        Boundaries.Bottom = -1
80        Erase bitfield
End Sub

Private Sub Class_Initialize()
10        ClearPoints
End Sub

Private Sub RegionFromPoints(add As Boolean)
          'that's the hard part...
          Dim i As Integer, j As Integer

          
          'ReDim isedge(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)
10        ReDim bitfield(Boundaries.Left - 1 To Boundaries.Right + 1, Boundaries.Top - 1 To Boundaries.Bottom + 1)
          'ReDim bitfield(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)
          
          'ReDim tmpedge(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)
          
          'Add contours
20        For i = 1 To PtCount - 1
30            bitfield(pts(i).X, pts(i).Y) = isedge
40            bitfield(pts(i - 1).X, pts(i - 1).Y) = isedge
              
50            If Abs(pts(i).X - pts(i - 1).X) > 1 Or Abs(pts(i).Y - pts(i - 1).Y) > 1 Then
60                Call EdgesFromLine(pts(i).X, pts(i).Y, pts(i - 1).X, pts(i - 1).Y)
70            End If
80        Next
          
          Dim pointsToCheck() As Coordinate
90        ReDim pointsToCheck(0)
100       pointsToCheck(0).X = Boundaries.Left - 1
110       pointsToCheck(0).Y = Boundaries.Top - 1
          Dim nrPoints As Long
120       nrPoints = 1
          
          
          Dim checkValue As Boolean
130       checkValue = False
          
          Dim tmparray() As Coordinate
          Dim tmpnr As Long
140       tmparray = pointsToCheck
150       tmpnr = nrPoints
160       Do
170           ReDim pointsToCheck(0)
180           nrPoints = 0
              
190           For i = 0 To tmpnr - 1
200               Call Regions(checkValue, tmparray(i).X, tmparray(i).Y, pointsToCheck, nrPoints)
210           Next
220           checkValue = Not checkValue
230           tmparray = pointsToCheck
240           tmpnr = nrPoints
              
250       Loop While UBound(pointsToCheck) > 0
          
          
          
          'Add to selection / Remove from selection
260       For j = Boundaries.Top To Boundaries.Bottom
270           For i = Boundaries.Left To Boundaries.Right
280               If bitfield(i, j) = setTrue Or bitfield(i, j) = checkedEdge Then
290                   If add Then
300                       Call parent.sel.AddSelectionArea(i, i, j, j, undoch, False)
310                   Else
320                       Call parent.sel.RemoveSelectionArea(i, i, j, j, undoch, False)
330                   End If
340               End If
350           Next
360       Next
End Sub



Private Sub EdgesFromLine(startx As Integer, startY As Integer, _
                         endX As Integer, endY As Integer)

          Dim InitialX As Integer
          Dim InitialY As Integer
          Dim FinalX As Integer
          Dim FinalY As Integer

          Dim i As Integer
          Dim j As Integer

10        InitialX = startx
20        InitialY = startY

30        FinalX = endX
40        FinalY = endY

          Dim correctStepX As Integer
          Dim correctStepY As Integer

          Dim Steep As Boolean
          Dim DeltaX As Integer, DeltaY As Integer, Delta As Integer
          Dim coord As Integer

50        DeltaX = Abs(endX - startx)
60        DeltaY = Abs(endY - startY)
          
          Dim StepX As Integer, StepY As Integer
          
70        If (FinalX - InitialX) > 0 Then
80            StepX = 1
90        Else
100           StepX = -1
110       End If
120       If (FinalY - InitialY) > 0 Then
130           StepY = 1
140       Else
150           StepY = -1
160       End If
          
170       If DeltaX < DeltaY Then
180           Steep = True
190           Swap InitialX, InitialY
200           Swap DeltaX, DeltaY
210           Swap StepX, StepY
220       Else
230           Steep = False
240       End If

250       Delta = (DeltaY + DeltaY) - DeltaX

260       For coord = 0 To DeltaX - 1
270           If Steep Then
280               bitfield(InitialY, InitialX) = isedge
290           Else
300               bitfield(InitialX, InitialY) = isedge
310           End If

320           While Delta >= 0
330               InitialY = InitialY + StepY
340               Delta = Delta - (DeltaX + DeltaX)
350           Wend
360           InitialX = InitialX + StepX
370           Delta = Delta + (DeltaY + DeltaY)
380       Next coord
          
End Sub

Private Sub Swap(X As Integer, Y As Integer)
          Dim tmp As Integer
10        tmp = X
20        X = Y
30        Y = tmp

End Sub



Private Sub Regions(val As Boolean, ByVal X As Integer, ByVal Y As Integer, ByRef pointsToCheck() As Coordinate, ByRef count As Long, Optional ByVal stack As Long = 50000)
          Dim px() As Long
          Dim py() As Long
          Dim Pstart As Long
          Dim Pstop As Long
          Dim Pcount As Long

          Dim Lx As Integer
          Dim Ly As Integer
          
          Dim Sloop As Long
10        ReDim px(0)
20        ReDim py(0)
          
30        Lx = X
40        Ly = Y

50        Pstart = 0
60        Pstop = 0
70        px(0) = Lx
80        py(0) = Ly

90        Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)

100       While Pstart <= Pstop

110           Lx = px(Pstart)
120           Ly = py(Pstart)
                      
130           If Not (Lx < Boundaries.Left - 1 Or Ly < Boundaries.Top - 1 Or Lx > Boundaries.Right + 1 Or Ly > Boundaries.Bottom + 1) Then

140               If bitfield(Lx, Ly) = notChecked Or bitfield(Lx, Ly) = isedge Then
150                   If (Lx < Boundaries.Left Or Ly < Boundaries.Top Or Lx > Boundaries.Right Or Ly > Boundaries.Bottom) Then
                          'we're just outside boundaries, add points
160                       Pcount = Pcount + 1
170                       bitfield(Lx, Ly) = setFalse
180                       Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)
190                   Else
200                       If bitfield(Lx, Ly) = notChecked Then
210                           If val Then
220                               bitfield(Lx, Ly) = setTrue
230                           Else
240                               bitfield(Lx, Ly) = setFalse
250                           End If
                              
260                           Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)
270                       Else
280                           bitfield(Lx, Ly) = checkedEdge
290                           Call AddPoint(Lx, Ly, pointsToCheck, count)
                              
300                       End If
                          
310                   End If
                                      
320               End If

330           End If
skiploop:
340           Pstart = Pstart + 1

350       Wend

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

110       ReDim Preserve px(Pstop + 4)
120       ReDim Preserve py(Pstop + 4)
130       Pstop = Pstop + 1
140       px(Pstop) = Lx + 1
150       py(Pstop) = Ly
160       Pstop = Pstop + 1
170       px(Pstop) = Lx
180       py(Pstop) = Ly + 1
190       Pstop = Pstop + 1
200       px(Pstop) = Lx - 1
210       py(Pstop) = Ly
220       Pstop = Pstop + 1
230       px(Pstop) = Lx
240       py(Pstop) = Ly - 1


          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub Class_Terminate()
10        Erase bitfield
20        Erase pts
30        Set parent = Nothing
40        Set undoch = Nothing
          
End Sub
