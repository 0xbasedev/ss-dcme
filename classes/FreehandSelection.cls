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
    isDrawing = c_isdrawing
End Property

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew
    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023

    If parent.sel.hasAlreadySelectedParts Then
        'check if we click inside the selection or outside, and with which button
        If Button = vbRightButton Then
            Set undoch = New Changes
            parent.undoredo.ResetRedo

            Call parent.sel.ApplySelection(undoch, True)

            Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_APPLY)

        ElseIf Button = vbLeftButton Then
            'If isInSelection(curtilex, curtiley) And _
             '    Shift <> 1 And Shift <> 2 Then
            If parent.sel.getIsInSelection(curtilex, curtiley) And ((parent.pastetype <> p_normal And parent.sel.getSelTile(curtilex, curtiley) <> 0) Or parent.pastetype = p_normal) And _
               Shift <> 1 And Shift <> 2 Then
                'we clicked in the selection, ready to move
                dragxcorrection = curtilex
                dragycorrection = curtiley
                previousX = curtilex
                previousY = curtiley
                movingselection = True
              
                  Call parent.sel.StartMove
            Else
                Set undoch = New Changes
                parent.undoredo.ResetRedo
                
                If Shift <> 1 And Shift <> 2 Then
                    Call parent.sel.ApplySelection(undoch, True)
                End If

                'we want to draw if we clicked outside of selection or
                ' if we use shift or ctrl

                'we want to draw again
                Call ClearPoints
                Call AddPoint(curtilex, curtiley, pts, PtCount)
                
                Boundaries.Right = curtilex
                Boundaries.Left = curtilex
                Boundaries.Bottom = curtiley
                Boundaries.Top = curtiley
                        
                previousX = curtilex
                previousY = curtiley
                c_isdrawing = True
                
                Call parent.UpdatePreview(True, False)

            End If
        End If
    Else
        Set undoch = New Changes
        parent.undoredo.ResetRedo

        'we want to draw if we clicked outside of selection or
        ' if we use shift or ctrl

        'we want to draw again
        Call ClearPoints
        Call AddPoint(curtilex, curtiley, pts, PtCount)
        
        Boundaries.Right = curtilex
        Boundaries.Left = curtilex
        Boundaries.Bottom = curtiley
        Boundaries.Top = curtiley
                
        previousX = curtilex
        previousY = curtiley
        c_isdrawing = True
        
        Call parent.UpdatePreview(True, False)
    End If

End Sub

Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    
    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew
    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023

    If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) And ((parent.pastetype <> p_normal And parent.sel.getSelTile(curtilex, curtiley) <> 0) Or parent.pastetype = p_normal) And _
       Shift <> 1 And Shift <> 2 Then
        parent.picPreview.MousePointer = 5
    Else
        parent.picPreview.MousePointer = 2
    End If
    
    If c_isdrawing Then
        If curtilex <> previousX Or curtiley <> previousY Then
            
            Call AddPoint(curtilex, curtiley, pts, PtCount)

            If curtilex > Boundaries.Right Then Boundaries.Right = curtilex
            If curtilex < Boundaries.Left Then Boundaries.Left = curtilex
            If curtiley > Boundaries.Bottom Then Boundaries.Bottom = curtiley
            If curtiley < Boundaries.Top Then Boundaries.Top = curtiley
            
            RedrawPoints True
            
            previousX = curtilex
            previousY = curtiley
        End If
        
        
    ElseIf movingselection Then
        If Button = vbLeftButton Then
            'we are dragging the selection around
            Call parent.sel.Move(curtilex - dragxcorrection, curtiley - dragycorrection)

            dragxcorrection = curtilex
            dragycorrection = curtiley
        End If
    End If
    
    
End Sub



Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    frmGeneral.IsBusy("FreehandSelection.MouseUp") = True
    
    If c_isdrawing Then
        Set undoch = New Changes
        Call parent.undoredo.ResetRedo
        
        If PtCount > 1 Then
            Call AddPoint(pts(0).X, pts(0).Y, pts, PtCount)
        End If
        RedrawPoints True
        
        Call RegionFromPoints(Not IsControl(Shift))
    
        c_isdrawing = False
        
        Call parent.undoredo.AddToUndo(undoch, "Freehand Selection")
        
        Call parent.UpdateLevel
    ElseIf movingselection Then
    
        Dim dx As Integer
        Dim dy As Integer
        Dim curtilex As Integer
        Dim curtiley As Integer
        
        'calculate tile coordinates
        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

        movingselection = False

        dx = curtilex - previousX
        dy = curtiley - previousY
        If dx <> 0 And dy <> 0 Then
            parent.undoredo.ResetRedo
            Set undoch = New Changes

              Call parent.sel.EndMove(undoch)

            Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_MOVE)
        End If
        
    End If
    
    frmGeneral.IsBusy("FreehandSelection.MouseUp") = False
    
End Sub

Sub RedrawPoints(Refresh As Boolean)
    Dim i As Integer
    Dim maplbx As Integer
    Dim maplby As Integer
    Dim maphbx As Integer
    Dim maphby As Integer

    maplbx = parent.Hscr.value \ parent.currenttilew
    maplby = parent.Vscr.value \ parent.currenttilew
    maphbx = maplbx + parent.picPreview.width \ parent.currenttilew
    maphby = maplby + parent.picPreview.height \ parent.currenttilew
    
    For i = 1 To PtCount - 1
      Call DrawLine(parent.hBufferDC, pts(i).X * parent.currenttilew - parent.Hscr.value + parent.currenttilew \ 2, _
                              pts(i).Y * parent.currenttilew - parent.Vscr.value + parent.currenttilew \ 2, _
                              pts(i - 1).X * parent.currenttilew - parent.Hscr.value + parent.currenttilew \ 2, _
                              pts(i - 1).Y * parent.currenttilew - parent.Vscr.value + parent.currenttilew \ 2, _
                              vbWhite)



'60            parent.picPreview.Line (pts(i).X * parent.currenttilew - parent.hScr.value + parent.currenttilew \ 2, pts(i).Y * parent.currenttilew - parent.vScr.value + parent.currenttilew \ 2)-(pts(i - 1).X * parent.currenttilew - parent.hScr.value + parent.currenttilew \ 2, pts(i - 1).Y * parent.currenttilew - parent.vScr.value + parent.currenttilew \ 2)
    Next

    If Refresh Then parent.UpdatePreview
End Sub

Private Sub AddPoint(X As Integer, Y As Integer, ByRef pts() As Coordinate, ByRef count As Long)
    If count > UBound(pts) Then
        ReDim Preserve pts(UBound(pts) + 500)
    End If
    
    pts(count).X = X
    pts(count).Y = Y
    
    count = count + 1
End Sub

Private Sub ClearPoints()
    PtCount = 0
    Erase pts
    ReDim pts(0)
    Boundaries.Right = -1
    Boundaries.Left = 1024
    Boundaries.Top = 1024
    Boundaries.Bottom = -1
    Erase bitfield
End Sub

Private Sub Class_Initialize()
    ClearPoints
End Sub

Private Sub RegionFromPoints(add As Boolean)
    'that's the hard part...
    Dim i As Integer, j As Integer

    
    'ReDim isedge(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)
    ReDim bitfield(Boundaries.Left - 1 To Boundaries.Right + 1, Boundaries.Top - 1 To Boundaries.Bottom + 1)
    'ReDim bitfield(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)
    
    'ReDim tmpedge(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)
    
    'Add contours
    For i = 1 To PtCount - 1
        bitfield(pts(i).X, pts(i).Y) = isedge
        bitfield(pts(i - 1).X, pts(i - 1).Y) = isedge
        
        If Abs(pts(i).X - pts(i - 1).X) > 1 Or Abs(pts(i).Y - pts(i - 1).Y) > 1 Then
            Call EdgesFromLine(pts(i).X, pts(i).Y, pts(i - 1).X, pts(i - 1).Y)
        End If
    Next
    
    Dim pointsToCheck() As Coordinate
    ReDim pointsToCheck(0)
    pointsToCheck(0).X = Boundaries.Left - 1
    pointsToCheck(0).Y = Boundaries.Top - 1
    Dim nrPoints As Long
    nrPoints = 1
    
    
    Dim checkValue As Boolean
    checkValue = False
    
    Dim tmparray() As Coordinate
    Dim tmpnr As Long
    tmparray = pointsToCheck
    tmpnr = nrPoints
    Do
        ReDim pointsToCheck(0)
        nrPoints = 0
        
        For i = 0 To tmpnr - 1
            Call Regions(checkValue, tmparray(i).X, tmparray(i).Y, pointsToCheck, nrPoints)
        Next
        checkValue = Not checkValue
        tmparray = pointsToCheck
        tmpnr = nrPoints
        
    Loop While UBound(pointsToCheck) > 0
    
    
    
    'Add to selection / Remove from selection
    For j = Boundaries.Top To Boundaries.Bottom
        For i = Boundaries.Left To Boundaries.Right
            If bitfield(i, j) = setTrue Or bitfield(i, j) = checkedEdge Then
                If add Then
                    Call parent.sel.AddSelectionArea(i, i, j, j, undoch, False)
                Else
                    Call parent.sel.RemoveSelectionArea(i, i, j, j, undoch, False)
                End If
            End If
        Next
    Next
End Sub



Private Sub EdgesFromLine(startx As Integer, startY As Integer, _
                         endX As Integer, endY As Integer)

    Dim InitialX As Integer
    Dim InitialY As Integer
    Dim FinalX As Integer
    Dim FinalY As Integer

    Dim i As Integer
    Dim j As Integer

    InitialX = startx
    InitialY = startY

    FinalX = endX
    FinalY = endY

    Dim correctStepX As Integer
    Dim correctStepY As Integer

    Dim Steep As Boolean
    Dim DeltaX As Integer, DeltaY As Integer, Delta As Integer
    Dim coord As Integer

    DeltaX = Abs(endX - startx)
    DeltaY = Abs(endY - startY)
    
    Dim StepX As Integer, StepY As Integer
    
    If (FinalX - InitialX) > 0 Then
        StepX = 1
    Else
        StepX = -1
    End If
    If (FinalY - InitialY) > 0 Then
        StepY = 1
    Else
        StepY = -1
    End If
    
    If DeltaX < DeltaY Then
        Steep = True
        Swap InitialX, InitialY
        Swap DeltaX, DeltaY
        Swap StepX, StepY
    Else
        Steep = False
    End If

    Delta = (DeltaY + DeltaY) - DeltaX

    For coord = 0 To DeltaX - 1
        If Steep Then
            bitfield(InitialY, InitialX) = isedge
        Else
            bitfield(InitialX, InitialY) = isedge
        End If

        While Delta >= 0
            InitialY = InitialY + StepY
            Delta = Delta - (DeltaX + DeltaX)
        Wend
        InitialX = InitialX + StepX
        Delta = Delta + (DeltaY + DeltaY)
    Next coord
    
End Sub

Private Sub Swap(X As Integer, Y As Integer)
    Dim tmp As Integer
    tmp = X
    X = Y
    Y = tmp

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
    ReDim px(0)
    ReDim py(0)
    
    Lx = X
    Ly = Y

    Pstart = 0
    Pstop = 0
    px(0) = Lx
    py(0) = Ly

    Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)

    While Pstart <= Pstop

        Lx = px(Pstart)
        Ly = py(Pstart)
                
        If Not (Lx < Boundaries.Left - 1 Or Ly < Boundaries.Top - 1 Or Lx > Boundaries.Right + 1 Or Ly > Boundaries.Bottom + 1) Then

            If bitfield(Lx, Ly) = notChecked Or bitfield(Lx, Ly) = isedge Then
                If (Lx < Boundaries.Left Or Ly < Boundaries.Top Or Lx > Boundaries.Right Or Ly > Boundaries.Bottom) Then
                    'we're just outside boundaries, add points
                    Pcount = Pcount + 1
                    bitfield(Lx, Ly) = setFalse
                    Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)
                Else
                    If bitfield(Lx, Ly) = notChecked Then
                        If val Then
                            bitfield(Lx, Ly) = setTrue
                        Else
                            bitfield(Lx, Ly) = setFalse
                        End If
                        
                        Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)
                    Else
                        bitfield(Lx, Ly) = checkedEdge
                        Call AddPoint(Lx, Ly, pointsToCheck, count)
                        
                    End If
                    
                End If
                                
            End If

        End If
skiploop:
        Pstart = Pstart + 1

    Wend

End Sub

Private Sub AddPoints(ByRef px() As Long, ByRef py() As Long, ByRef Pstart As Long, ByRef Pstop As Long, ByRef Lx As Integer, ByRef Ly As Integer, ByRef Sloop As Long, ByRef stack As Long)

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

    ReDim Preserve px(Pstop + 4)
    ReDim Preserve py(Pstop + 4)
    Pstop = Pstop + 1
    px(Pstop) = Lx + 1
    py(Pstop) = Ly
    Pstop = Pstop + 1
    px(Pstop) = Lx
    py(Pstop) = Ly + 1
    Pstop = Pstop + 1
    px(Pstop) = Lx - 1
    py(Pstop) = Ly
    Pstop = Pstop + 1
    px(Pstop) = Lx
    py(Pstop) = Ly - 1


    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub Class_Terminate()
    Erase bitfield
    Erase pts
    Set parent = Nothing
    Set undoch = Nothing
    
End Sub
