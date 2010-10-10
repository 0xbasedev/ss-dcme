VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim firstX As Integer
Dim firstY As Integer


Dim dragxcorrection As Integer
Dim dragycorrection As Integer

Enum selectionstate
    notdrawing
    drawing
    Append
End Enum

Dim c_selstate As selectionstate

Dim seltilestartx As Integer
Dim seltilestarty As Integer
Dim seltileendx As Integer
Dim seltileendy As Integer

Dim parent As frmMain

Dim undoch As Changes

Dim seltiles() As Integer
Dim isInSelection As boolArray

Public hasAlreadySelectedParts As Boolean

Dim selectionchanged As Boolean

Dim movingselection As Boolean

Dim boundaries As area
Dim tempboundaries As area

'Boundaries before the move starts
Dim moveboundaries As area


Dim cursormoved As Boolean

Private Type SpecialObject
    X As Integer
    Y As Integer
    tilenr As Integer
End Type

Private Sub Class_Initialize()
    boundaries.Top = 32767
    boundaries.Left = 32767
    boundaries.Right = 0
    boundaries.Bottom = 0
    tempboundaries.Top = 32767
    tempboundaries.Left = 32767
    tempboundaries.Right = 0
    tempboundaries.Bottom = 0

    ReDim seltiles(1023, 1023)
    
    Set isInSelection = New boolArray
    Call isInSelection.InitArray(1024, 1024, 0)
'    ReDim isInSelection(1023, 1023)

End Sub

Sub SetParent(ByRef main As frmMain)
    Set parent = main
End Sub

Function getIsInSelection(i As Integer, j As Integer) As Boolean
    getIsInSelection = isInSelection.value(i, j)
End Function

Function getBitField() As boolArray
    Set getBitField = isInSelection
End Function

'Function isedge(i As Integer, j As Integer) As Boolean
'
'    If parent.pastetype = p_normal Then
'        If isInSelection.value(i - 1, j) And _
'           isInSelection.value(i, j - 1) And _
'           isInSelection.value(i + 1, j) And _
'           isInSelection.value(i, j + 1) Then
'            isedge = False
'        Else
'            isedge = True
'        End If
'    Else
'        If seltiles(i, j) Then
'            'with other pastings its only in selection when isinselection AND not seltiles = 0
'            If isInSelection.value(i - 1, j) And seltiles(i - 1, j) And _
'               isInSelection.value(i, j - 1) And seltiles(i, j - 1) And _
'               isInSelection.value(i + 1, j) And seltiles(i + 1, j) And _
'               isInSelection.value(i, j + 1) And seltiles(i, j + 1) Then
'                isedge = False
'            Else
'                isedge = True
'            End If
'        Else
'            isedge = False
'        End If
'    End If
'
'End Function

Friend Function getBoundaries() As area
    getBoundaries = boundaries
End Function

Public Property Get selstate() As selectionstate
    selstate = c_selstate
End Property

Public Property Let selstate(ByVal newselstate As selectionstate)
    c_selstate = newselstate
End Property

Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer

    'calculate tile coordinates
    On Error GoTo MouseDown_Error

    curtilex = (parent.hScr.value + X) \ parent.currenttilew
    curtiley = (parent.vScr.value + Y) \ parent.currenttilew

    cursormoved = False

    If curtilex > 1023 Then curtilex = 1023
    If curtilex < 0 Then curtilex = 0
    If curtiley > 1023 Then curtiley = 1023
    If curtiley < 0 Then curtiley = 0

    firstX = curtilex
    firstY = curtiley

    If selstate = Append Then
        'check if we click inside the selection or outside, and with which button
        If Button = vbRightButton Then
            Set undoch = New Changes
            parent.undoredo.ResetRedo

            Call ApplySelection(undoch, True)

            Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_APPLY)

        ElseIf Button = vbLeftButton Then
            'If isInSelection(curtilex, curtiley) And _
             '    Shift <> 1 And Shift <> 2 Then
            If isInSelection.value(curtilex, curtiley) And _
               Shift <> 1 And Shift <> 2 Then
                'we clicked in the selection, ready to move
                dragxcorrection = curtilex
                dragycorrection = curtiley

                movingselection = True
                
                Call StartMove
                
            ElseIf Shift = 1 Or Shift = 2 Then
                'we want to draw if we clicked outside of selection or
                ' if we use shift or ctrl

                'we want to draw again
                selstate = drawing
                Call initDrawing(X, Y)
            Else
                selstate = notdrawing
                Set undoch = New Changes
                parent.undoredo.ResetRedo

                Call ApplySelection(undoch, True)
                Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_APPLY)
            End If
        End If
    End If

    If selstate = notdrawing And Button = vbLeftButton Then
        'we want to start drawing
        Call initDrawing(X, Y)
    End If

    'Check if we tried to cancel selection by right-clicking
    If selstate = drawing And Button = vbRightButton Then
        'Clear selection
        StopSelecting
    End If

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "Selection.MouseDown"
End Sub

Sub StopSelecting()
    Call MouseUp(vbRightButton, 1, 10, 10)
    selstate = Append
End Sub

Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer

'    On Error GoTo MouseMove_Error

    cursormoved = True
    
    If Button = vbRightButton Then Exit Sub

    'calculate tile coordinates
    curtilex = (parent.hScr.value + X) \ parent.currenttilew
    curtiley = (parent.vScr.value + Y) \ parent.currenttilew

    If curtilex > 1023 Then curtilex = 1023
    If curtilex < 0 Then curtilex = 0
    If curtiley > 1023 Then curtiley = 1023
    If curtiley < 0 Then curtiley = 0

    If hasAlreadySelectedParts And isInSelection.value(curtilex, curtiley) And _
       Shift <> 1 And Shift <> 2 Then
        parent.picpreview.MousePointer = 5
    Else
        parent.picpreview.MousePointer = 2
    End If

    If SharedVar.MouseDown = 0 Then Exit Sub

    If selstate = drawing Then

        'Swap first/current values if needed
        If firstX < curtilex Then
            seltilestartx = firstX
            seltileendx = curtilex
        Else
            seltilestartx = curtilex
            seltileendx = firstX
        End If

        If firstY < curtiley Then
            seltilestarty = firstY
            seltileendy = curtiley
        Else
            seltilestarty = curtiley
            seltileendy = firstY
        End If


        'update the drawing rectangle
        parent.shpdraw.Left = (seltilestartx * parent.currenttilew) - parent.hScr.value
        parent.shpdraw.Top = (seltilestarty * parent.currenttilew) - parent.vScr.value
        parent.shpdraw.width = (seltileendx - seltilestartx + 1) * parent.currenttilew
        parent.shpdraw.height = (seltileendy - seltilestarty + 1) * parent.currenttilew

    ElseIf selstate = Append Then

        If Button = vbLeftButton Then
            'we are dragging the selection around
            Call Move(curtilex - dragxcorrection, curtiley - dragycorrection)

            dragxcorrection = curtilex
            dragycorrection = curtiley
        End If

    End If

'    On Error GoTo 0
'    Exit Sub
'
'MouseMove_Error:
'    HandleError Err, "Selection.MouseMove"
End Sub

Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not bDEBUG Then On Error GoTo MouseUp_Error

    frmGeneral.IsBusy("Selection.MouseUp") = True
    
    If parent.magnifier.usingpixels Then
        'stupid selection on stupid pixel level has wrong offsets
        'because of the +2, but thats needed in non-pixel level
        'to have the correct tile star
        seltilestartx = (parent.hScr.value + parent.shpdraw.Left) \ parent.currenttilew
        seltilestarty = (parent.vScr.value + parent.shpdraw.Top) \ parent.currenttilew
        seltileendx = seltilestartx + (parent.shpdraw.width / (parent.currenttilew)) - 1
        seltileendy = seltilestarty + (parent.shpdraw.height / (parent.currenttilew)) - 1
    Else
        seltilestartx = (parent.hScr.value + parent.shpdraw.Left + 2) \ parent.currenttilew
        seltilestarty = (parent.vScr.value + parent.shpdraw.Top + 2) \ parent.currenttilew
        seltileendx = seltilestartx + (parent.shpdraw.width / (parent.currenttilew)) - 1
        seltileendy = seltilestarty + (parent.shpdraw.height / (parent.currenttilew)) - 1
    End If

    parent.shpdraw.visible = False
    parent.shpdraw.Left = 0
    parent.shpdraw.Top = 0
    parent.shpdraw.width = 1
    parent.shpdraw.height = 1

    If movingselection And cursormoved Then
        'we're done moving selection
        Dim curtilex As Integer
        Dim curtiley As Integer
        Dim dx As Integer
        Dim dy As Integer

        'calculate tile coordinates
        curtilex = (parent.hScr.value + X) \ parent.currenttilew
        curtiley = (parent.vScr.value + Y) \ parent.currenttilew

        movingselection = False

        dx = curtilex - firstX
        dy = curtiley - firstY
        If dx <> 0 And dy <> 0 Then
            parent.undoredo.ResetRedo
            Set undoch = New Changes

            Call EndMove(undoch)

            Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_MOVE)
        End If

    ElseIf Button = vbLeftButton And selstate = drawing And cursormoved Then

        If Shift = 2 Then    'ctrl
            parent.undoredo.ResetRedo
            Set undoch = New Changes

            If RemoveSelectionArea(seltilestartx, seltileendx, seltilestarty, seltileendy, undoch) Then
                Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_REMOVEAREA)
            End If

        ElseIf Shift = 1 Then    ' shift
            parent.undoredo.ResetRedo
            Set undoch = New Changes

            If AddSelectionArea(seltilestartx, seltileendx, seltilestarty, seltileendy, undoch) Then
                Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_ADDAREA)
            End If
        Else
            If hasAlreadySelectedParts Then
                parent.undoredo.ResetRedo
                Set undoch = New Changes

                If AddSelectionArea(seltilestartx, seltileendx, seltilestarty, seltileendy, undoch) Then
                    Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_ADDAREA)
                End If
            Else
                parent.undoredo.ResetRedo
                Set undoch = New Changes

                Call ApplySelection(undoch, False)
                Call AddSelectionArea(seltilestartx, seltileendx, seltilestarty, seltileendy, undoch)

                Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_APPLY_AND_ADD)
            End If
        End If
        selstate = Append



    Else
        'We have right-clicked, so we're not adding anything to the selection
    End If


    If parent.magnifier.usingpixels Then
        parent.UpdateLevel
    Else
        parent.UpdatePreview
    End If

'    frmGeneral.LockBusyState = False
    frmGeneral.IsBusy("Selection.MouseUp") = False
    
    On Error GoTo 0
    Exit Sub

MouseUp_Error:
    
    frmGeneral.IsBusy("Selection.MouseUp") = False
    HandleError Err, "Selection.MouseUp " & Button & " " & Shift & " (" & X & "," & Y & ")"
End Sub

Sub initDrawing(X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer


    On Error GoTo initDrawing_Error

    c_selstate = drawing

    'determine starting point of the selection
    curtilex = (parent.hScr.value + X) \ parent.currenttilew
    curtiley = (parent.vScr.value + Y) \ parent.currenttilew

    If curtilex > 1023 Then curtilex = 1023
    If curtilex < 0 Then curtilex = 0
    If curtiley > 1023 Then curtiley = 1023
    If curtiley < 0 Then curtiley = 0

    parent.shpdraw.Left = (curtilex * parent.currenttilew) - parent.hScr.value
    parent.shpdraw.Top = (curtiley * parent.currenttilew) - parent.vScr.value
    parent.shpdraw.width = 1
    parent.shpdraw.height = 1
    parent.shpdraw.visible = True

    On Error GoTo 0
    Exit Sub

initDrawing_Error:
    HandleError Err, "Selection.initDrawing"
End Sub

'Function areThereTilesInSelection() As Boolean
'    Dim i As Integer, j As Integer
'    For j = 0 To 1023
'        For i = 0 To 1023
'            If isInSelection.value(i, j) Then
'                areThereTilesInSelection = True
'                Exit Function
'            End If
'        Next
'    Next
'    areThereTilesInSelection = False
'End Function



Function AddSelectionArea(seltilestartx As Integer, seltileendx As Integer, seltilestarty As Integer, seltileendy As Integer, undoch As Changes, Optional Refresh As Boolean = True) As Boolean
    Dim tick As Long
    tick = GetTickCount
    
    Dim i As Integer
    Dim j As Integer
    Dim tmp As Integer

    '    Dim t1 As Long
    '    t1 = GetTickCount

    On Error GoTo AddSelectionArea_Error

    Dim hasAdded As Boolean
    hasAdded = False

    If boundaries.Left > seltilestartx Then boundaries.Left = seltilestartx
    If boundaries.Right < seltileendx Then boundaries.Right = seltileendx
    If boundaries.Top > seltilestarty Then boundaries.Top = seltilestarty
    If boundaries.Bottom < seltileendy Then boundaries.Bottom = seltileendy
    tempboundaries = boundaries

    For j = seltilestarty To seltileendy
        For i = seltilestartx To seltileendx
            If Not isInSelection.value(i, j) Then
                tmp = parent.getTile(i, j)
                If tmp <> 0 Then
                    If tmp = 217 Or tmp = 219 Or tmp = 220 Or tmp < 0 Then
                        'special tile found, add it completly
                        Call AddObjectToSelection(i, j, undoch)

                        If tempboundaries.Left < boundaries.Left Then boundaries.Left = tempboundaries.Left
                        If tempboundaries.Right > boundaries.Right Then boundaries.Right = tempboundaries.Right
                        If tempboundaries.Top < boundaries.Top Then boundaries.Top = tempboundaries.Top
                        If tempboundaries.Bottom > boundaries.Bottom Then boundaries.Bottom = tempboundaries.Bottom

                    Else
                        Call AddTileToSelection(i, j, undoch)

                        '                        '  remove the tiles from the level
                        '                        Call parent.setTile(i, j, 0, undoch, True)
                        '                        Call parent.UpdateLevelTile(i, j, False)

                    End If


                Else
                    Call AddTileToSelection(i, j, undoch)
                End If


                '                'copy all the tiles to the selection array
                '                Call setSelTile(i, j, tmp, undoch)
                '                'flag that they are on the selection
                '                isInSelection(i, j) = True
                hasAdded = True
            End If
        Next
    Next



    If hasAdded Then
        selstate = Append
        hasAlreadySelectedParts = True
    End If

    If Refresh And hasAdded Then
        parent.pic1024selection.Refresh
        frmGeneral.UpdateToolBarButtons
        parent.UpdateLevel
    End If

    AddSelectionArea = hasAdded

    frmGeneral.Label6.Caption = "AddSelectionArea: " & GetTickCount - tick
    
    On Error GoTo 0
    Exit Function

AddSelectionArea_Error:
    HandleError Err, "Selection.AddSelectionArea"
End Function

Function RemoveSelectionArea(seltilestartx As Integer, seltileendx As Integer, seltilestarty As Integer, seltileendy As Integer, undoch As Changes, Optional Refresh As Boolean = True) As Boolean
    Dim tick As Long
    tick = GetTickCount
    
    Dim i As Integer
    Dim j As Integer

    Dim foundtiles As Boolean
    Dim removedtiles As Boolean

    On Error GoTo RemoveSelectionArea_Error

    If boundaries.Left > boundaries.Right Or boundaries.Top > boundaries.Bottom Then Exit Function

    'If the area covers the whole selection, just apply it
    If seltilestartx <= boundaries.Left And seltileendx >= boundaries.Right _
     And seltilestarty <= boundaries.Top And seltileendy >= boundaries.Bottom Then
        Call ApplySelection(undoch, Refresh)
        RemoveSelectionArea = True

    'If the area does not cover in part the boundaries, don't bother
    ElseIf seltileendx < boundaries.Left Or seltilestartx > boundaries.Right Or _
     seltileendy < boundaries.Top Or seltilestarty > boundaries.Bottom Then
        Exit Function
    End If
    
    'Don't go searching the selected area if it's not even part of the boundaries
'    If seltilestartx < Boundaries.Left Then seltilestartx = Boundaries.Left
'    If seltileendx > Boundaries.Right Then seltileendx = Boundaries.Right
'    If seltilestarty < Boundaries.Top Then seltilestarty = Boundaries.Top
'    If seltileendy > Boundaries.Bottom Then seltileendy = Boundaries.Bottom
    
    tempboundaries = boundaries
    foundtiles = False
    removedtiles = False

    'check for special objects being removed

    For j = seltilestarty To seltileendy
        For i = seltilestartx To seltileendx
            If isInSelection.value(i, j) Then
                Dim shouldbedropped As Boolean
                shouldbedropped = ((parent.pastetype = p_normal) Or _
                                   (parent.pastetype = p_trans And seltiles(i, j) <> 0) Or _
                                   (parent.pastetype = p_under And parent.getTile(i, j) = 0 And AreaClearForObject(parent, i, j, seltiles(i, j))))

                If shouldbedropped Then
                    'drop it

                    If TileIsSpecial(seltiles(i, j)) Then
                        Call DropObjectFromSelection(i, j, undoch)
                    Else
                        Call DropTileFromSelection(i, j, undoch)
                    End If

                    

                Else
                    'We still need to make the inselection False
                    Call undoch.AddTileChange(SelDelete, i, j, 0)

                    isInSelection.value(i, j) = False
                End If
                
                
                removedtiles = True
                '
                '                If isSelTileSpecial(Me, i, j) Then
                '                    'special tile found, removing it completly
                '                    Call RemoveObjectSelectionArea(i, j, undoch)
                '
                '                Else
                '                    If isTileSpecial(parent, i, j) And seltiles(i, j) <> 0 Then
                '                        Call SearchAndDestroyObject(parent, i, j, undoch, False)
                '                    End If
                '                    Call DropTileFromSelection(i, j, undoch)
                ''                    Call parent.setTile(i, j, seltiles(i, j), undoch, True)
                ''                    Call parent.UpdateLevelTile(i, j, False)
                ''                    'remove the tiles from the selection
                ''                    Call setSelTile(i, j, 0, undoch, False)
                ''                    'flag that they are on the selection
                ''                    isInSelection(i, j) = False
                '                End If



            End If
        Next
    Next

'    frmGeneral.pictemp.BackColor = vbWhite
'    frmGeneral.pictemp.width = seltileendx - seltilestartx + 1
'    frmGeneral.pictemp.height = seltileendy - seltilestarty + 1
'    BitBlt parent.pic1024selection.hdc, seltilestartx, seltilestarty, seltileendx - seltilestartx + 1, seltileendy - seltilestarty + 1, frmGeneral.pictemp.hdc, 0, 0, vbSrcCopy
'    frmGeneral.pictemp.BackColor = vbBlack
    
    
    
    ' check if we need to shrink the left side
    'If seltilestartx <= tempboundaries.Left Then
    If removedtiles Then
        j = 0: i = 0
        For j = tempboundaries.Left To boundaries.Right
            For i = boundaries.Top To boundaries.Bottom
                If isInSelection.value(j, i) Then
                    foundtiles = True
                    GoTo skipLeft
                End If
            Next
            tempboundaries.Left = tempboundaries.Left + 1
        Next
        'End If
skipLeft:

        ' check if we need to shrink the top side
        'If seltilestarty <= tempboundaries.Top Then
        j = 0: i = 0
        For j = tempboundaries.Top To boundaries.Bottom
            For i = boundaries.Left To boundaries.Right
                If isInSelection.value(i, j) Then
                    foundtiles = True
                    GoTo skipTop
                End If
            Next
            tempboundaries.Top = tempboundaries.Top + 1
        Next
        'End If
skipTop:

        ' check if we need to shrink the right side
        'If seltileendx >= tempboundaries.Right Then
        j = 0: i = 0
        For j = tempboundaries.Right To boundaries.Left Step -1
            For i = boundaries.Top To boundaries.Bottom
                If isInSelection.value(j, i) Then
                    foundtiles = True
                    GoTo skipRight
                End If
            Next
            tempboundaries.Right = tempboundaries.Right - 1
        Next
        'End If
skipRight:

        ' check if we need to shrink the right side
        'If seltileendy >= tempboundaries.Bottom Then
        j = 0: i = 0
        For j = tempboundaries.Bottom To boundaries.Top Step -1
            For i = boundaries.Left To boundaries.Right
                If isInSelection.value(i, j) Then
                    foundtiles = True
                    GoTo skipBottom
                End If
            Next
            tempboundaries.Bottom = tempboundaries.Bottom - 1
        Next
        'End If
skipBottom:

    End If

    boundaries = tempboundaries

    hasAlreadySelectedParts = foundtiles
        
    If Refresh Then
        'check if there are any tiles left in selection
        If Not hasAlreadySelectedParts Then
            parent.pic1024selection.Cls
        End If

        frmGeneral.UpdateToolBarButtons
        Call parent.RedrawSelectionArea((seltilestartx - 1) * TILEW, (seltilestarty - 1) * TILEH, (seltileendx + 1) * TILEW, (seltileendy + 1) * TILEH, False)
        Call parent.RedrawTileLayer(True)
'        parent.UpdateLevel
    End If

    RemoveSelectionArea = removedtiles

    frmGeneral.Label6.Caption = "RemoveSelectionArea: " & GetTickCount - tick
    
    On Error GoTo 0
    Exit Function

RemoveSelectionArea_Error:
    HandleError Err, "Selection.RemoveSelectionArea"
End Function

Function RemoveAllSelectionArea(undoch As Changes, Optional Refresh As Boolean = True) As Boolean
    'optimized function of removeselectionarea (when all has to be cleared)
    Dim shouldbedropped As Boolean
    Dim oldtile As Integer, newtile As Integer
    
    Dim removedtiles As Boolean
    
    Dim i As Integer, j As Integer
    
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            If isInSelection.value(i, j) Then
            
                newtile = seltiles(i, j)
                oldtile = parent.getTile(i, j)
                
                If parent.pastetype = p_normal Then
                    shouldbedropped = True
                ElseIf parent.pastetype = p_trans And newtile <> 0 Then
                    shouldbedropped = True
                ElseIf parent.pastetype = p_under And oldtile = 0 And AreaClearForObject(parent, i, j, newtile) Then
                    shouldbedropped = True
                Else
                    shouldbedropped = False
                End If
                
                If shouldbedropped Then
                    If TileIsSpecial(newtile) Then
                        Call DropObjectFromSelection(i, j, undoch)
                    Else
                    
                        Call parent.setTile(i, j, newtile, Nothing, False)
                        'Call parent.UpdateLevelTile(x, y, False, False)
                        'false, false, meaning just setting pixel
                        'Call setPixel(parent.pic1024.hdc, i, j, TilePixelColor(seltiles(i, j)))
                        Call parent.cpic1024.setPixelLong(i, j, TilePixelColor(newtile))
                        
                        Call undoch.AddTileChange(SelDrop, i, j, oldtile)
                    
                        'doesn't matter, killing selection anyway
                        'isInSelection(x, y) = False
                        'Call setSelTile(x, y, 0, Nothing, True, False)
                    End If
                    removedtiles = True
                End If
            End If
        Next
    Next
    
    RemoveAllSelectionArea = removedtiles
End Function

Sub ApplySelection(undoch As Changes, Optional Refresh As Boolean = True)
    
    On Error GoTo ApplySelection_Error

    'apply selection is nothing more than removeselection but on the whole selection
    'array, rather than a part of it
    Dim tick As Long
    tick = GetTickCount
    
    Dim i As Integer
    Dim j As Integer

    frmGeneral.IsBusy("Selection.ApplySelection") = True
    
    'use optimized form
    Call RemoveAllSelectionArea(undoch, False)
    'Call RemoveSelectionArea(Boundaries.Left, Boundaries.Right, Boundaries.Top, Boundaries.Bottom, undoch, False)

    '    For j = Boundaries.Top To Boundaries.Bottom
    '        For i = Boundaries.Left To Boundaries.Right
    '            If isInSelection(i, j) Then
    '                'copy all the tiles to the level array
    '
    '                'check for draw type
    '                If Not seltiles(i, j) < 0 And _
                     '                   ((parent.pastetype = p_normal) Or _
                     '                    (parent.pastetype = p_trans And seltiles(i, j) <> 0) Or _
                     '                    (parent.pastetype = p_under And parent.getTile(i, j) = 0 And AreaClearForObject(parent, i, j, seltiles(i, j)))) Then
    '
    '                    Call DropTileFromSelection(i, j, undoch)
    '                End If
    '            End If
    '        Next
    '    Next


    hasAlreadySelectedParts = False
    parent.pic1024selection.Cls

    
    'Call parent.UpdateLevel

    'reset selection state
    c_selstate = notdrawing
    selectionchanged = False

    boundaries.Top = 32767
    boundaries.Left = 32767
    boundaries.Right = 0
    boundaries.Bottom = 0


    'clear possible tile nr's stored in unused selection tiles
    Call DeleteSelection(Nothing, False, False)

    frmGeneral.UpdateToolBarButtons

    If Refresh Then parent.UpdateLevel

    frmGeneral.IsBusy("Selection.ApplySelection") = False

    frmGeneral.Label6.Caption = "ApplySelection: " & GetTickCount - tick
    
    On Error GoTo 0
    Exit Sub

ApplySelection_Error:
    frmGeneral.IsBusy("Selection.ApplySelection") = False
    HandleError Err, "Selection.ApplySelection"
End Sub

Sub CopySelection()
'set the array into the clipboard array
    frmGeneral.IsBusy("Selection.CopySelection") = True
    
    Call frmGeneral.clipboard.SetClipboardData(seltiles, isInSelection, boundaries)


    frmGeneral.IsBusy("Selection.CopySelection") = False
End Sub

Sub CutSelection()
'set the array into the clipboard array
'remove the selection
    frmGeneral.IsBusy("Selection.CutSelection") = True
    
    CopySelection

    parent.undoredo.ResetRedo
    Set undoch = New Changes

    Call DeleteSelection(undoch, True, True)

    Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_CLEAR)
    
    frmGeneral.IsBusy("Selection.CutSelection") = False
End Sub

Sub DeleteSelection(undoch As Changes, Optional Refresh As Boolean = True, Optional appendundo As Boolean = False)
'clear the selection, but don't copy the tiles from the selection to the level
    frmGeneral.IsBusy("Selection.DeleteSelection") = True
    
    If appendundo Then

        Dim i As Integer, j As Integer
        
        For i = boundaries.Left To boundaries.Right
            For j = boundaries.Top To boundaries.Bottom
                If isInSelection.value(i, j) Then
                    Call DeleteSelectionTile(i, j, undoch)
                End If
            Next
        Next

    Else
        Dim ln As Long
        ln = (CLng(1024) * 1024 - 1) * 2
        ZeroMemory ByVal VarPtr(seltiles(0, 0)), ln
        
        Call isInSelection.FillArray(0)
'        ZeroMemory ByVal VarPtr(isInSelection(0, 0)), ln
    End If

    hasAlreadySelectedParts = False

    parent.pic1024selection.Cls

    frmGeneral.UpdateToolBarButtons

    If Refresh Then
        parent.UpdateLevel
    End If
    
    frmGeneral.IsBusy("Selection.DeleteSelection") = False
    
End Sub

Sub DrawCornerLinesOn(i As Integer, j As Integer, lbx As Integer, lby As Integer, curtilew As Integer, ByRef layer As clsDisplayLayer)

'    If parent.pastetype <> p_normal And seltiles(i, j) = 0 Then
'        Exit Sub
'    End If
    Dim neighbors As Byte
    Dim hDC As Long
    hDC = layer.hDC
    
    neighbors = isInSelection.GetNeighborhood(i, j, False)
    
    If curtilew = 1 Then
        'Using pixels
        'tile left of it is not in selection

        
        If (neighbors And Neighbor.N_Left) = Neighbor.N_Left Then Call SetPixel(hDC, lbx - 1, lby, vbYellow)
        
            'tile top of it is not in selection
        If (neighbors And Neighbor.N_Top) = Neighbor.N_Top Then Call SetPixel(hDC, lbx, lby - 1, vbYellow)
    
            'tile right of it is not in selection
        If (neighbors And Neighbor.N_Right) = Neighbor.N_Right Then Call SetPixel(hDC, lbx + 1, lby, vbYellow)
    
            'tile bottom of it is not in selection
        If (neighbors And Neighbor.N_Bottom) = Neighbor.N_Bottom Then Call SetPixel(hDC, lbx, lby + 1, vbYellow)
    
    
    Else
    
    
        'tile left of it is not in selection
        If (neighbors And Neighbor.N_Left) = Neighbor.N_Left Then Call DrawLine(hDC, lbx, lby, lbx, lby + curtilew, vbWhite)
        
            'tile top of it is not in selection
        If (neighbors And Neighbor.N_Top) = Neighbor.N_Top Then Call DrawLine(hDC, lbx, lby, lbx + curtilew, lby, vbWhite)
    
            'tile right of it is not in selection
        If (neighbors And Neighbor.N_Right) = Neighbor.N_Right Then Call DrawLine(hDC, lbx + curtilew - 1, lby, lbx + curtilew - 1, lby + curtilew, vbWhite)
    
            'tile bottom of it is not in selection
        If (neighbors And Neighbor.N_Bottom) = Neighbor.N_Bottom Then Call DrawLine(hDC, lbx, lby + curtilew - 1, lbx + curtilew, lby + curtilew - 1, vbWhite)
    End If




End Sub



Friend Sub SetSelectionData(tiles() As Integer, ByRef bitfield As boolArray, ByRef Bounds As area, Optional newundo As Boolean = True, Optional undocomment As String = UNDO_SELECTION_PASTE)
    On Error GoTo SetSelectionData_Error

    frmGeneral.IsBusy("Selection.SetSelectionData") = True
    
    If newundo Then
        parent.undoredo.ResetRedo
        Set undoch = New Changes
    End If

    If hasAlreadySelectedParts Then
        Call ApplySelection(undoch, False)
    End If

    Call undoch.AddSelPaste

    seltiles = tiles
    Call isInSelection.CopyFromBitfield(bitfield)
    boundaries = Bounds

    hasAlreadySelectedParts = True
    selstate = Append

    parent.pic1024selection.Cls
    Dim i As Integer
    Dim j As Integer
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            If isInSelection.value(i, j) Then
                Call SetPixel(parent.pic1024selection.hDC, i, j, TilePixelColor(seltiles(i, j)))
            End If
        Next
    Next
    parent.pic1024selection.Refresh

    Call CenterSelection(True, False, undoch)
    
'    'move the selection boundaries to the top left of the picpreview,
'    'but if a part will be cut off then move it until no parts are being cut off
'    Dim lbx As Integer    'lowestval for x
'    Dim hbx As Integer    'highestval for x
'    Dim lby As Integer    'lowestval for y
'    Dim hby As Integer    'highestval for y
'    lbx = parent.Hscr.value \ parent.currenttilew
'    hbx = (parent.Hscr.value + parent.picPreview.width) \ parent.currenttilew
'    lby = parent.Vscr.value \ parent.currenttilew
'    hby = (parent.Vscr.value + parent.picPreview.height) \ parent.currenttilew
'
'    Dim dx As Integer
'    dx = lbx - Boundaries.Left
'    If (Boundaries.Right + dx) > 1023 Then
'        dx = dx - ((Boundaries.Right + dx) - 1024)
'    End If
'
'    Dim dy As Integer
'    dy = lby - Boundaries.Top
'    If (Boundaries.Bottom + dy) > 1023 Then
'        dy = dy - ((Boundaries.Bottom + dy) - 1024)
'    End If
'
'    If dx <> 0 Or dy <> 0 Then
'        Call Move(dx, dy, False, True)
'        Call undoch.AddSelMove(dx, dy)
'    End If

    If newundo Then
        Call parent.undoredo.AddToUndo(undoch, undocomment)
    End If

    frmGeneral.UpdateToolBarButtons
    
    Call parent.UpdateLevel

    frmGeneral.IsBusy("Selection.SetSelectionData") = False
    
    On Error GoTo 0
    Exit Sub

SetSelectionData_Error:

    frmGeneral.IsBusy("Selection.SetSelectionData") = False
    
    HandleError Err, "Selection.SetSelectionData"
End Sub

Sub Move(dx As Integer, dy As Integer, Optional Refresh As Boolean = True, Optional movePixels As Boolean = True)
'    On Error GoTo Move_Error

    Dim tick As Long
    tick = GetTickCount

    If boundaries.Left + dx < 0 Then dx = 0 - boundaries.Left
    If boundaries.Right + dx > 1023 Then dx = 1023 - boundaries.Right
    If boundaries.Top + dy < 0 Then dy = 0 - boundaries.Top
    If boundaries.Bottom + dy > 1023 Then dy = 1023 - boundaries.Bottom

    If dx = 0 And dy = 0 Then
        'no move
        Exit Sub
    End If
    
'
'    If Boundaries.Left + dx < 0 Or Boundaries.Top + dy < 0 Or _
'       Boundaries.Right + dx > 1023 Or Boundaries.Bottom + dy > 1023 Then
        
'        Exit Sub
'    End If

    Dim newseltiles(1023, 1023) As Integer
'    Dim newisInSelection(1023, 1023) As Boolean


    
    Dim i As Integer, j As Integer
    
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
        '    If isInSelection(i, j) Then
                newseltiles(i + dx, j + dy) = seltiles(i, j)
'                newisInSelection(i + dx, j + dY) = isInSelection.value(i, j)
         '   End If
        Next
    Next
    
    'apply the new array to the old one
    seltiles = newseltiles
'    Call isInSelection.SetFromBooleanArray(newisInSelection, 1024, 1024)

    Call isInSelection.MoveRectangle(boundaries.Left, boundaries.Right, boundaries.Top, boundaries.Bottom, dx, dy, True)
    
    
    boundaries.Left = boundaries.Left + dx
    boundaries.Top = boundaries.Top + dy
    boundaries.Right = boundaries.Right + dx
    boundaries.Bottom = boundaries.Bottom + dy
    
    'don't update pic1024selection if not on pixel level, it slows down alot
    If movePixels Then
        parent.picseltemp.Cls
        Call BitBlt(parent.picseltemp.hDC, 0, 0, 1024, 1024, parent.pic1024selection.hDC, 0, 0, vbSrcCopy)
        parent.pic1024selection.Cls
        'parent.picseltemp.refresh
        Call BitBlt(parent.pic1024selection.hDC, dx, dy, 1024, 1024, parent.picseltemp.hDC, 0, 0, vbSrcCopy)
        'parent.pic1024selection.refresh
    End If

    'remember that the selection has been moved
    selectionchanged = True


    If Refresh Then
        Call parent.RedrawSelection(True)
    End If
    
    frmGeneral.Label6.Caption = "selmove:" & GetTickCount - tick
    
'    On Error GoTo 0
'    Exit Sub
'
'Move_Error:
'    HandleError Err, "Selection.Move"
End Sub

Sub setSelTile(ByRef X As Integer, ByRef Y As Integer, ByRef val As Integer, undoch As Changes, Optional DrawPixel As Boolean = True, Optional appendundo As Boolean = True)
    If seltiles(X, Y) <> val Then
        Dim old As Integer
        old = seltiles(X, Y)
        If appendundo Then
            Call undoch.AddTileChange(SelTileChange, X, Y, old)
        End If

        seltiles(X, Y) = val
    End If

    '   isInSelection(i, j) = True
    If DrawPixel Then
        Call SetPixel(parent.pic1024selection.hDC, X, Y, TilePixelColor(val))
    End If

    '    If tile(tileX, tileY) <> value Then
    '        'only call undo when tile is changed
    '        Dim old As Integer
    '        old = tile(tileX, tileY)
    '        If appendundo Then
    '            Call undoch.AddTileChange(MapTileChange, tileX, tileY, old)
    '        End If
    '        tile(tileX, tileY) = value
    '    End If
End Sub

Sub setSelTileAndisIn(ByRef X As Integer, ByRef Y As Integer, ByRef tilenr As Integer, ByRef isin As Boolean, undoch As Changes)
    If isin Then
        If isInSelection.value(X, Y) Then
            'tile was simply changed
            Call setSelTile(X, Y, tilenr, undoch)
        Else
            'tile is being created in selection
            Call CreateTileInSelection(X, Y, tilenr, undoch)
        End If
    Else
        If isInSelection.value(X, Y) Then
            'tile is removed from selection
            Call DeleteSelectionTile(X, Y, undoch)
        Else
            'it wasn't in selection, and will not be either; nothing to do
        End If
    End If
End Sub

Function getSelTile(ByRef i As Integer, ByRef j As Integer) As Integer
    getSelTile = seltiles(i, j)
End Function

Sub SwitchOrReplace(src As Integer, dest As Integer, replace As Boolean, Redraw As Boolean)
'for each isinselection(i,j) between the boundaries replace or switch
    Dim i As Integer
    Dim j As Integer
    Dim m As Long
    Dim n As Long
    Dim p As Long

    On Error GoTo SwitchOrReplace_Error

    n = 0
    m = 0
    p = 0

    Dim issrcvalid(255) As Boolean
    Dim isdestvalid(255) As Boolean
    Dim srcwalltile As Boolean
    Dim destwalltile As Boolean
    Dim srcwallset As Integer
    Dim destwallset As Integer
    Dim tmp As Integer
    Dim curtile As Integer

    parent.undoredo.ResetRedo
    Set undoch = New Changes

    If src >= 259 Then
        srcwallset = src - 259
        srcwalltile = True
        For i = 0 To 15
            issrcvalid(parent.walltiles.getWallTile(src - 259, i)) = True
        Next
    End If

    If dest >= 259 Then
        destwallset = dest - 259
        destwalltile = True
        For i = 0 To 15
            isdestvalid(parent.walltiles.getWallTile(dest - 259, i)) = True
        Next
    End If

    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            If isInSelection.value(i, j) Then
                If seltiles(i, j) < 0 Then
                    curtile = seltiles(i, j) \ -100
                Else
                    curtile = seltiles(i, j)
                End If
                If curtile = src Or (srcwalltile And issrcvalid(curtile)) Then
                    If destwalltile Then
                        tmp = parent.walltiles.ReplaceSelWithWalltile(i, j, destwallset, src, srcwalltile, srcwallset, replace)
                    Else
                        tmp = dest
                    End If
                    Call setSelTile(i, j, tmp, undoch)
                    n = n + 1
                    'if the tile is a dest tile and we switch
                    'then we replace it with the src tile
                ElseIf Not replace And (curtile = dest Or (destwalltile And isdestvalid(curtile))) Then
                    If srcwalltile Then
                        parent.walltiles.curwall = srcwallset
                        tmp = parent.walltiles.ReplaceSelWithWalltile(i, j, srcwallset, dest, destwalltile, destwallset, replace)
                    Else
                        tmp = src
                    End If
                    Call setSelTile(i, j, tmp, undoch)
                    m = m + 1
                ElseIf replace And Redraw And destwalltile And isdestvalid(curtile) Then
                    tmp = parent.walltiles.ReplaceSelWithWalltile(i, j, destwallset, src, srcwalltile, srcwallset, True)

                    Call setSelTile(i, j, tmp, undoch)
                    p = p + 1
                End If
            End If
        Next
    Next

    If replace Then
        MessageBox n & " tiles replaced" & vbCrLf & p & " walltiles redrawn", vbOKOnly + vbInformation, "Tiles replaced"
    Else
        MessageBox n & " tiles A -> B" & vbCrLf & m & " tiles A <- B", vbOKOnly + vbInformation, "Tiles switched"
    End If

    Call parent.undoredo.AddToUndo(undoch, IIf(replace, UNDO_SELECTION_REPLACE, UNDO_SELECTION_SWITCH))

    parent.UpdateLevel

    selectionchanged = True

    On Error GoTo 0
    Exit Sub

SwitchOrReplace_Error:
    HandleError Err, "Selection.SwitchOrReplace"
End Sub

Sub Mirror(undoch As Changes)
'mirror entire selection array

    Dim i As Integer
    Dim j As Integer

    Dim tmp As Integer
    Dim tmpisin As Boolean

    On Error GoTo Mirror_Error

    Call undoch.AddSelFlip(True)

    Dim tmpobjects() As SpecialObject
    ReDim tmpobjects(10)
    Dim tmpobjectsCount As Long
    tmpobjectsCount = 0

    'first off, calculate the new positions of any objects
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            tmp = seltiles(i, j)

            If tmp = 217 Or tmp = 219 Or tmp = 220 Then
                If tmpobjectsCount > UBound(tmpobjects) Then
                    ReDim Preserve tmpobjects(tmpobjectsCount + 20)
                End If

                tmpobjects(tmpobjectsCount).X = boundaries.Left + boundaries.Right - (i + GetMaxSizeOfObject(tmp))
                tmpobjects(tmpobjectsCount).Y = j
                tmpobjects(tmpobjectsCount).tilenr = tmp

                Call DestroySelObject(Me, i, j, Nothing, , , False)
                tmpobjectsCount = tmpobjectsCount + 1

            End If
        Next
    Next

    'flip the selection array
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Left + ((boundaries.Right - boundaries.Left) \ 2)

            tmp = seltiles(i, j)
            tmpisin = isInSelection.value(i, j)

            isInSelection.value(i, j) = isInSelection.value(boundaries.Right - (i - boundaries.Left), j)
            isInSelection.value(boundaries.Right - (i - boundaries.Left), j) = tmpisin

            'ignore special objects
            Call setSelTile(i, j, seltiles(boundaries.Right - (i - boundaries.Left), j), Nothing, True, False)
            Call setSelTile(boundaries.Right - (i - boundaries.Left), j, tmp, Nothing, True, False)

        Next
    Next

    'now that it's flipped we restore the objects
    If tmpobjectsCount > 0 Then
        For i = 0 To tmpobjectsCount - 1
            Call setSelObject(Me, tmpobjects(i).tilenr, tmpobjects(i).X, tmpobjects(i).Y, Nothing, False, False)
        Next
    End If

    selectionchanged = True

    Call parent.UpdateLevel

    On Error GoTo 0
    Exit Sub

Mirror_Error:
    HandleError Err, "Selection.Mirror"
End Sub

Sub Flip(undoch As Changes)
'mirror entire selection array

    Dim i As Integer
    Dim j As Integer

    Dim tmp As Integer
    Dim tmpisin As Boolean

    On Error GoTo Flip_Error

    Call undoch.AddSelFlip(False)

    Dim tmpobjects() As SpecialObject
    ReDim tmpobjects(10)
    Dim tmpobjectsCount As Long
    tmpobjectsCount = 0

    'first off, calculate the new positions of any objects
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            tmp = seltiles(i, j)

            If tmp = 217 Or tmp = 219 Or tmp = 220 Then
                If tmpobjectsCount > UBound(tmpobjects) Then
                    ReDim Preserve tmpobjects(tmpobjectsCount + 20)
                End If

                tmpobjects(tmpobjectsCount).X = i
                tmpobjects(tmpobjectsCount).Y = boundaries.Top + boundaries.Bottom - (j + GetMaxSizeOfObject(tmp))
                tmpobjects(tmpobjectsCount).tilenr = tmp

                Call DestroySelObject(Me, i, j, Nothing, , , False)
                tmpobjectsCount = tmpobjectsCount + 1
            End If
        Next
    Next

    'flip the selection array
    For i = boundaries.Left To boundaries.Right
        For j = boundaries.Top To boundaries.Top + ((boundaries.Bottom - boundaries.Top) \ 2)

            tmp = seltiles(i, j)
            tmpisin = isInSelection.value(i, j)

            '            Call setSelTileAndisIn(i, j, seltiles(i, Boundaries.Bottom - (j - Boundaries.Top)), isInSelection(i, Boundaries.Bottom - (j - Boundaries.Top)), undoch)
            '            Call setSelTileAndisIn(i, Boundaries.Bottom - (j - Boundaries.Top), tmp, tmpisin, undoch)
            isInSelection.value(i, j) = isInSelection.value(i, boundaries.Bottom - (j - boundaries.Top))
            isInSelection.value(i, boundaries.Bottom - (j - boundaries.Top)) = tmpisin

            'ignore special objects
            Call setSelTile(i, j, seltiles(i, boundaries.Bottom - (j - boundaries.Top)), Nothing, True, False)
            Call setSelTile(i, boundaries.Bottom - (j - boundaries.Top), tmp, Nothing, True, False)
        Next
    Next

    'now that it's flipped we restore the objects
    If tmpobjectsCount > 0 Then
        For i = 0 To tmpobjectsCount - 1
            Call setSelObject(Me, tmpobjects(i).tilenr, tmpobjects(i).X, tmpobjects(i).Y, Nothing, False, False)
        Next
    End If

    selectionchanged = True

    Call parent.UpdateLevel

    On Error GoTo 0
    Exit Sub

Flip_Error:
    HandleError Err, "Selection.Flip"
End Sub


Sub RotateCW(undoch As Changes)
    Dim i As Integer
    Dim j As Integer

    Dim selctrx As Integer
    Dim selctry As Integer
    Dim newboundaries As area

    Dim tmp As Integer

    Call undoch.AddSelRotate(90)

    'New parameters of selection
    selctrx = (boundaries.Right - boundaries.Left) \ 2
    selctry = (boundaries.Bottom - boundaries.Top) \ 2
    newboundaries.Left = (selctrx + boundaries.Left) - (selctry)
    newboundaries.Top = (selctry + boundaries.Top) - (selctrx)
    newboundaries.Bottom = newboundaries.Top + (boundaries.Right - boundaries.Left)
    newboundaries.Right = newboundaries.Left + (boundaries.Bottom - boundaries.Top)

    Dim tmptile() As Integer
    ReDim tmptile(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)

    Dim tmpisin() As Boolean
    ReDim tmpisin(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)


    'Make sure new parameters aren't out of bounds
    If newboundaries.Left < 0 Then
        newboundaries.Left = 0
        newboundaries.Right = newboundaries.Left + (boundaries.Bottom - boundaries.Top)
    ElseIf newboundaries.Right > 1023 Then
        newboundaries.Left = newboundaries.Left + (1023 - newboundaries.Right)
        newboundaries.Right = 1023
    End If

    If newboundaries.Top < 0 Then
        newboundaries.Top = 0
        newboundaries.Bottom = newboundaries.Top + (boundaries.Right - boundaries.Left)
    ElseIf newboundaries.Bottom > 1023 Then
        newboundaries.Top = newboundaries.Top + (1023 - newboundaries.Bottom)
        newboundaries.Bottom = 1023
    End If

    Dim tmpobjects() As SpecialObject
    ReDim tmpobjects(10)
    Dim tmpobjectsCount As Long
    tmpobjectsCount = 0

    'first off, calculate the new positions of any objects
    For i = boundaries.Left To boundaries.Right
        For j = boundaries.Top To boundaries.Bottom
            tmp = seltiles(i, j)
            tmpisin(i, j) = isInSelection.value(i, j)
            If isInSelection.value(i, j) And (tmp = 217 Or tmp = 219 Or tmp = 220) Then
                '''Special tile found, calculating new coordinates
                If tmpobjectsCount > UBound(tmpobjects) Then
                    ReDim Preserve tmpobjects(tmpobjectsCount + 20)
                End If

                tmpobjects(tmpobjectsCount).X = newboundaries.Right - ((j - boundaries.Top) + GetMaxSizeOfObject(tmp))
                tmpobjects(tmpobjectsCount).Y = newboundaries.Top + (i - boundaries.Left)
                tmpobjects(tmpobjectsCount).tilenr = tmp

                Call DestroySelObject(Me, i, j, Nothing, , , False)
                tmpobjectsCount = tmpobjectsCount + 1
            Else

                '''Store tile in current selection
                tmptile(i, j) = seltiles(i, j)
            End If

            'Clean that tile off the selection
            Call setinselection(i, j, False)
            Call setSelTile(i, j, 0, Nothing, True, False)
        Next
    Next



    'Retile the new selection with corresponding tiles
    For j = newboundaries.Top To newboundaries.Bottom
        For i = newboundaries.Left To newboundaries.Right
            isInSelection.value(i, j) = tmpisin(boundaries.Left + (j - newboundaries.Top), boundaries.Top + (newboundaries.Right - i))
            If tmptile(boundaries.Left + (j - newboundaries.Top), boundaries.Top + (newboundaries.Right - i)) >= 0 Then
                Call setSelTile(i, j, tmptile(boundaries.Left + (j - newboundaries.Top), boundaries.Top + (newboundaries.Right - i)), Nothing, True, False)
            End If

        Next
    Next

    'now that it's flipped we restore the objects
    If tmpobjectsCount > 0 Then
        For i = 0 To tmpobjectsCount - 1
            Call setSelObject(Me, tmpobjects(i).tilenr, tmpobjects(i).X, tmpobjects(i).Y, Nothing, False, False)
        Next
    End If

    'Redefine selection with new parameters
    boundaries.Left = newboundaries.Left
    boundaries.Top = newboundaries.Top
    boundaries.Right = newboundaries.Right
    boundaries.Bottom = newboundaries.Bottom


    selectionchanged = True

    Call parent.UpdateLevel
End Sub

Sub RotateCCW(undoch As Changes)
    Dim i As Integer
    Dim j As Integer

    Dim selctrx As Integer
    Dim selctry As Integer

    Dim newboundaries As area

    Dim tmp As Integer

    Call undoch.AddSelRotate(270)

    'New parameters of selection
    selctrx = (boundaries.Right - boundaries.Left) \ 2
    selctry = (boundaries.Bottom - boundaries.Top) \ 2
    newboundaries.Left = (selctrx - (selctry - boundaries.Left))
    newboundaries.Top = (selctry - (selctrx - boundaries.Top))

    newboundaries.Right = newboundaries.Left + (boundaries.Bottom - boundaries.Top)
    newboundaries.Bottom = newboundaries.Top + (boundaries.Right - boundaries.Left)

    Dim tmptile() As Integer
    ReDim tmptile(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)

    Dim tmpisin() As Boolean
    ReDim tmpisin(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)




    'Make sure new parameters aren't out of bounds
    If newboundaries.Left < 0 Then
        newboundaries.Left = 0
        newboundaries.Right = newboundaries.Left + (boundaries.Bottom - boundaries.Top)
    ElseIf newboundaries.Right > 1023 Then
        newboundaries.Left = newboundaries.Left + (1023 - newboundaries.Right)
        newboundaries.Right = 1023
    End If

    If newboundaries.Top < 0 Then
        newboundaries.Top = 0
        newboundaries.Bottom = newboundaries.Top + (boundaries.Right - boundaries.Left)
    ElseIf newboundaries.Bottom > 1023 Then
        newboundaries.Top = newboundaries.Top + (1023 - newboundaries.Bottom)
        newboundaries.Bottom = 1023
    End If

    Dim tmpobjects() As SpecialObject
    ReDim tmpobjects(10)
    Dim tmpobjectsCount As Long
    tmpobjectsCount = 0

    'first off, calculate the new positions of any objects
    For i = boundaries.Left To boundaries.Right
        For j = boundaries.Top To boundaries.Bottom
            tmp = seltiles(i, j)
            tmpisin(i, j) = isInSelection.value(i, j)
            If isInSelection.value(i, j) And (tmp = 217 Or tmp = 219 Or tmp = 220) Then
                '''Special tile found, calculating new coordinates
                If tmpobjectsCount > UBound(tmpobjects) Then
                    ReDim Preserve tmpobjects(tmpobjectsCount + 20)
                End If

                tmpobjects(tmpobjectsCount).X = newboundaries.Left + (j - boundaries.Top)
                tmpobjects(tmpobjectsCount).Y = newboundaries.Bottom - ((i - boundaries.Left) + GetMaxSizeOfObject(tmp))
                tmpobjects(tmpobjectsCount).tilenr = tmp

                Call DestroySelObject(Me, i, j, Nothing, , , False)
                tmpobjectsCount = tmpobjectsCount + 1

            Else

                '''Store tile in current selection
                tmptile(i, j) = seltiles(i, j)
            End If

            'Clean that tile off the selection
            Call setinselection(i, j, False)
            Call setSelTile(i, j, 0, Nothing, True, False)
        Next
    Next

    'Retile the new selection with corresponding tiles
    For j = newboundaries.Top To newboundaries.Bottom
        For i = newboundaries.Left To newboundaries.Right
            isInSelection.value(i, j) = tmpisin(boundaries.Left + (newboundaries.Bottom - j), boundaries.Top + (i - newboundaries.Left))
            If tmptile(boundaries.Left + (newboundaries.Bottom - j), boundaries.Top + (i - newboundaries.Left)) >= 0 Then
                Call setSelTile(i, j, tmptile(boundaries.Left + (newboundaries.Bottom - j), boundaries.Top + (i - newboundaries.Left)), Nothing, True, False)
            End If

        Next
    Next

    'now that it's flipped we restore the objects
    If tmpobjectsCount > 0 Then
        For i = 0 To tmpobjectsCount - 1
            Call setSelObject(Me, tmpobjects(i).tilenr, tmpobjects(i).X, tmpobjects(i).Y, Nothing, False, False)
        Next
    End If


    'Redefine selection with new parameters
    boundaries.Left = newboundaries.Left
    boundaries.Top = newboundaries.Top
    boundaries.Right = newboundaries.Right
    boundaries.Bottom = newboundaries.Bottom



    selectionchanged = True
    Call parent.UpdateLevel
End Sub

Sub Rotate180(undoch As Changes)
    Dim i As Integer
    Dim j As Integer
    Dim tmp As Integer

    Dim tmptile() As Integer
    ReDim tmptile(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)

    Dim tmpisin() As Boolean
    ReDim tmpisin(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)

    Call undoch.AddSelRotate(180)

    Dim tmpobjects() As SpecialObject
    ReDim tmpobjects(10)
    Dim tmpobjectsCount As Long
    tmpobjectsCount = 0

    'first off, calculate the new positions of any objects
    For i = boundaries.Left To boundaries.Right
        For j = boundaries.Top To boundaries.Bottom
            tmp = seltiles(i, j)
            tmpisin(i, j) = isInSelection.value(i, j)
            If isInSelection.value(i, j) And (tmp = 217 Or tmp = 219 Or tmp = 220) Then
                'Special object found
                If tmpobjectsCount > UBound(tmpobjects) Then
                    ReDim Preserve tmpobjects(tmpobjectsCount + 20)
                End If

                tmpobjects(tmpobjectsCount).X = boundaries.Left + boundaries.Right - (i + GetMaxSizeOfObject(tmp))
                tmpobjects(tmpobjectsCount).Y = boundaries.Top + boundaries.Bottom - (j + GetMaxSizeOfObject(tmp))
                tmpobjects(tmpobjectsCount).tilenr = tmp

                Call DestroySelObject(Me, i, j, Nothing, , , False)
                tmpobjectsCount = tmpobjectsCount + 1
            Else

                'Store tile
                tmptile(i, j) = seltiles(i, j)
            End If
        Next
    Next

    'now that the objects are temp gone we can flip the image
    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            Call setSelTile(i, j, tmptile(boundaries.Right - (i - boundaries.Left), boundaries.Bottom - (j - boundaries.Top)), Nothing, True, False)
            isInSelection.value(i, j) = tmpisin(boundaries.Right - (i - boundaries.Left), boundaries.Bottom - (j - boundaries.Top))
        Next
    Next

    'now that it's flipped we restore the objects
    If tmpobjectsCount > 0 Then
        For i = 0 To tmpobjectsCount - 1
            Call setSelObject(Me, tmpobjects(i).tilenr, tmpobjects(i).X, tmpobjects(i).Y, Nothing, False, False)
        Next
    End If

    selectionchanged = True

    Call parent.UpdateLevel

End Sub

Sub RotateFree(angle As Double, undoch As Changes)
    Dim i As Integer
    Dim j As Integer

    Dim m As Single
    Dim n As Single

    Dim tmp As Integer

    Dim tmpdestx As Integer
    Dim tmpdesty As Integer
    Dim tmporigx As Integer
    Dim tmporigy As Integer
    Dim tmporigxsng As Single
    Dim tmporigysng As Single

    Dim tmptile(1023, 1023) As Integer
    'Erase tmptile
    'ReDim tmptile(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)

    Dim tmpisin(1023, 1023) As Boolean
    'Erase tmpisin
    'ReDim tmpisin(boundaries.Left To boundaries.Right, boundaries.Top To boundaries.Bottom)

    Dim tmpobjects() As SpecialObject
    ReDim tmpobjects(10)
    Dim tmpobjectsCount As Long
    tmpobjectsCount = 0

    Dim cosangle As Double
    Dim sinangle As Double
    cosangle = Cos(angle)
    sinangle = Sin(angle)

    Dim prompted As Boolean

    'first off, calculate the new positions of any objects

    tempboundaries = boundaries

    tmporigy = boundaries.Top
    tmporigysng = boundaries.Top
    For n = -(boundaries.Bottom - boundaries.Top) \ 2 To (boundaries.Bottom - boundaries.Top + 0.5) \ 2 Step 1!
        tmporigx = boundaries.Left
        tmporigxsng = boundaries.Left
        For m = -(boundaries.Right - boundaries.Left) \ 2 To (boundaries.Right - boundaries.Left + 0.5) \ 2 Step 1!

            tmp = seltiles(tmporigx, tmporigy)
            tmpisin(tmporigx, tmporigy) = isInSelection.value(tmporigx, tmporigy)

            If CSng(tmporigx) = tmporigxsng And CSng(tmporigy) = tmporigysng Then
                If isInSelection.value(tmporigx, tmporigy) And (tmp = 217 Or tmp = 219 Or tmp = 220) Then
                    '''Special tile found, warn user of distortion
                    If Not prompted Then
                        If MessageBox("Special objects were found in your selection. The result may be distorted because special objects cannot be rotated. Do you want to continue?", vbYesNo + vbQuestion, "Rotating Special Objects") = vbNo Then
                            Exit Sub
                        Else
                            prompted = True
                        End If
                    End If

                    If tmpobjectsCount > UBound(tmpobjects) Then
                        ReDim Preserve tmpobjects(tmpobjectsCount + 20)
                    End If

                    'calculate new coordinates
                    tmpdestx = Int(cosangle * (m + (GetMaxSizeOfObject(tmp) + 1) / 2) + sinangle * (n + (GetMaxSizeOfObject(tmp) + 1) / 2)) + boundaries.Left + (boundaries.Right - boundaries.Left) \ 2 - (GetMaxSizeOfObject(tmp) + 1) / 2
                    tmpdesty = Int(-sinangle * (m + (GetMaxSizeOfObject(tmp) + 1) / 2) + cosangle * (n + (GetMaxSizeOfObject(tmp) + 1) / 2)) + boundaries.Top + (boundaries.Bottom - boundaries.Top) \ 2 - (GetMaxSizeOfObject(tmp) + 1) / 2

                    tmpobjects(tmpobjectsCount).X = tmpdestx
                    tmpobjects(tmpobjectsCount).Y = tmpdesty
                    tmpobjects(tmpobjectsCount).tilenr = tmp
                    tmpobjectsCount = tmpobjectsCount + 1

                    If tmpisin(tmporigx, tmporigy) Then
                        If tmpdestx < tempboundaries.Left Then tempboundaries.Left = tmpdestx
                        If tmpdestx + GetMaxSizeOfObject(tmp) > tempboundaries.Right Then tempboundaries.Right = tmpdestx + GetMaxSizeOfObject(tmp)
                        If tmpdesty < tempboundaries.Top Then tempboundaries.Top = tmpdesty
                        If tmpdesty + GetMaxSizeOfObject(tmp) > tempboundaries.Bottom Then tempboundaries.Bottom = tmpdesty + GetMaxSizeOfObject(tmp)
                    End If


                    Call DeleteSelectionTile(tmporigx, tmporigy, undoch)

                ElseIf isInSelection.value(tmporigx, tmporigy) And tmp >= 0 Then
                    'Store tile
                    tmptile(tmporigx, tmporigy) = tmp
                    Call DeleteSelectionTile(tmporigx, tmporigy, undoch)

                ElseIf isInSelection.value(tmporigx, tmporigy) Then
                    Call DeleteSelectionTile(tmporigx, tmporigy, undoch)

                End If
            End If

            tmporigxsng = tmporigxsng + 1
            tmporigx = Int(tmporigxsng)
        Next
        tmporigysng = tmporigysng + 1
        tmporigy = Int(tmporigysng)
    Next

    'clear objects
    '    If tmpobjectsCount > 0 Then
    '        For i = 0 To tmpobjectsCount - 1
    '            Call SearchAndDestroySelObject(Me, tmpobjects(i).X, tmpobjects(i).Y, undoch, True, True)
    '        Next
    '    End If
    '#########################
    '
    '    For i = Boundaries.Left To Boundaries.Right
    '        For j = Boundaries.Top To Boundaries.Bottom
    '            tmp = seltiles(i, j)
    '            tmpisin(i, j) = isinselection.value(i, j)
    '            If isinselection.value(i, j) And (tmp = 217 Or tmp = 219 Or tmp = 220) Then
    '                If Not prompted Then
    '                    If messagebox("Special objects were found in your selection. The result may be distorted because special objects cannot be rotated. Do you want to continue?", vbYesNo + vbQuestion, "Rotating Special Objects") = vbNo Then
    '                        Exit Sub
    '                    Else
    '                        prompted = True
    '                    End If
    '                End If
    '
    '                tmpdestx = Int(cosangle * m + sinangle * n) + Boundaries.Left + (Boundaries.Right - Boundaries.Left) \ 2
    '                tmpdesty = Int(-sinangle * m + cosangle * n) + Boundaries.Top + (Boundaries.Bottom - Boundaries.Top) \ 2
    '
    '                '              tmpobjects(Int(Cos(angle) * i + Sin(angle) * j), Int(-Sin(angle) * i + Cos(angle) * j)) = tmp
    '
    '            Else
    '                'Store tile
    '                tmptile(i, j) = seltiles(i, j)
    '            End If
    '
    '            Call DeleteSelectionTile(i, j, undoch)
    '        Next
    '    Next

    'now that the objects are temp gone we can flip the image
    '    For j = boundaries.Top To boundaries.Bottom
    '        For i = boundaries.Left To boundaries.Right
    '            tmpdestx = Int(Cos(angle) * i + Sin(angle) * j)
    '            tmpdesty = Int(-Sin(angle) * i + Cos(angle) * j)
    '            Call setSelTile(tmpdestx, tmpdesty, tmptile(i, j))
    '            isinselection.value(tmpdestx, tmpdesty) = tmpisin(i, j)
    '        Next
    '    Next



    tmporigy = boundaries.Top
    tmporigysng = boundaries.Top
    For n = -(boundaries.Bottom - boundaries.Top) \ 2 To (boundaries.Bottom - boundaries.Top + 0.5) \ 2 Step 0.25
        tmporigx = boundaries.Left
        tmporigxsng = boundaries.Left
        For m = -(boundaries.Right - boundaries.Left) \ 2 To (boundaries.Right - boundaries.Left + 0.5) \ 2 Step 0.25
            tmpdestx = Int(cosangle * m + sinangle * n) + boundaries.Left + (boundaries.Right - boundaries.Left) \ 2
            tmpdesty = Int(-sinangle * m + cosangle * n) + boundaries.Top + (boundaries.Bottom - boundaries.Top) \ 2

            'Call setSelTileAndisIn(tmpdestx, tmpdesty, tmptile(tmporigx, tmporigy), tmpisin(tmporigx, tmporigy), undoch)
            Call setSelTile(tmpdestx, tmpdesty, tmptile(tmporigx, tmporigy), Nothing, True, False)
            isInSelection.value(tmpdestx, tmpdesty) = tmpisin(tmporigx, tmporigy)

            If tmpisin(tmporigx, tmporigy) Then
                If tmpdestx < tempboundaries.Left Then tempboundaries.Left = tmpdestx
                If tmpdestx > tempboundaries.Right Then tempboundaries.Right = tmpdestx
                If tmpdesty < tempboundaries.Top Then tempboundaries.Top = tmpdesty
                If tmpdesty > tempboundaries.Bottom Then tempboundaries.Bottom = tmpdesty
            End If

            tmporigxsng = tmporigxsng + 0.25
            tmporigx = Int(tmporigxsng)
        Next
        tmporigysng = tmporigysng + 0.25
        tmporigy = Int(tmporigysng)
    Next



    'now that it's flipped we restore the objects
    If tmpobjectsCount > 0 Then
        For i = 0 To tmpobjectsCount - 1
            Call setSelObject(Me, tmpobjects(i).tilenr, tmpobjects(i).X, tmpobjects(i).Y, Nothing, True, False)
        Next
    End If

    'scan the new area and record undo actions
    For i = tempboundaries.Left To tempboundaries.Right
        For j = tempboundaries.Top To tempboundaries.Bottom
            If isInSelection.value(i, j) Then
                Call undoch.AddTileChange(SelNew, i, j, seltiles(i, j))
            End If
        Next
    Next

    boundaries = tempboundaries

    selectionchanged = True

    Call parent.UpdateLevel

End Sub


Sub TextToMap(Text() As Integer, width As Integer, height As Integer, Optional undocomment As String = UNDO_TEXTTOMAP)
    Dim bnds As area
    Dim selectiontiles() As Integer
    Dim bitfield As New boolArray
    
    On Error GoTo TexTToMap_Error

    frmGeneral.IsBusy("Selection.TextToMap") = True
    
    ReDim selectiontiles(1023, 1023)
    
    Call bitfield.InitArray(1024, 1024, 0)
    
    bnds.Left = 0
    bnds.Top = 0
    bnds.Right = width - 1
    bnds.Bottom = height - 1

    Call bitfield.SetRectangle(0, width - 1, 0, height - 1, 1)
    
    Dim i As Integer
    Dim j As Integer
    For j = 0 To height - 1
        For i = 0 To width - 1
            selectiontiles(i, j) = Text(i, j)
'            bitfield.value(i, j) = True
        Next
    Next

    Call SetSelectionData(selectiontiles, bitfield, bnds, True, undocomment)
    
    frmGeneral.IsBusy("Selection.TextToMap") = False
    
    On Error GoTo 0
    Exit Sub

TexTToMap_Error:
    frmGeneral.IsBusy("Selection.TextToMap") = False
    
    HandleError Err, "Selection.TextToMap"
End Sub

Sub PicToMap(Pic() As Integer, width As Integer, height As Integer)
    Call TextToMap(Pic(), width, height, UNDO_PICTOMAP)
End Sub

'Sub RemoveObjectSelectionArea(X As Integer, Y As Integer, undoch As changes)
'    Dim i As Integer
'    Dim j As Integer
'    Dim coord() As Integer
'    Dim coordX As Integer
'    Dim coordY As Integer
'    Dim size As Integer
'
'    On Error GoTo RemoveObjectSelectionArea_Error
'
'    coord = SearchSelObject(Me, X, Y)
'    coordX = coord(0)
'    coordY = coord(1)
'    size = GetMaxSizeOfObject(seltiles(coordX, coordY))
'
'    For j = coordY To coordY + size
'        For i = coordX To coordX + size
'            If isInSelection(i, j) Then
'                Call DropTileFromSelection(i, j, undoch)
''                'copy all the tiles to the level array
''                Call parent.setTile(i, j, seltiles(i, j), undoch, True)
''                Call parent.UpdateLevelTile(i, j, False)
''                'remove the tiles from the selection
''                Call setSelTile(i, j, 0, False)
''                'flag that they are on the selection
''                isInSelection(i, j) = False
'                Call SetPixel(parent.pic1024selection.hDC, i, j, vbWhite)
'            End If
'        Next
'    Next
'
'    parent.mapchanged = True
'    On Error GoTo 0
'    Exit Sub
'
'RemoveObjectSelectionArea_Error:
'
'    messagebox "Error " & err.Number & " (" & err.Description & ") in procedure RemoveObjectSelectionArea of Class Module Selection"
'
'End Sub
'Sub AddObjectSelectionArea(X As Integer, Y As Integer, undoch As changes)
'    Dim i As Integer
'    Dim j As Integer
'    Dim coord() As Integer
'    Dim coordX As Integer
'    Dim coordY As Integer
'    Dim size As Integer
'
'    On Error GoTo AddObjectSelectionArea_Error
'
'    ReDim coord(1) As Integer
'    coord = SearchObject(parent, X, Y)
'    coordX = coord(0)
'    coordY = coord(1)
'
'    Call setSelObject(Me, parent.getTile(coordX, coordY), coordX, coordY, undoch, True)
'
'    size = GetMaxSizeOfObject(parent.getTile(coordX, coordY))
'
'    Call DestroyObject(parent, coordX, coordY, undoch, False)
'
'
''
''    For j = coordY To coordY + size
''        For i = coordX To coordX + size
''
''            Call AddTileToSelection(i, j, undoch)
''        Next
''    Next
'
'    'Store coordinates to adjust boundaries if needed
'    tempboundaries.Left = coordX
'    tempboundaries.Right = coordX + size
'    tempboundaries.Top = coordY
'    tempboundaries.Bottom = coordY + size
'
'    parent.mapchanged = True
'    On Error GoTo 0
'    Exit Sub
'
'AddObjectSelectionArea_Error:
'
'    messagebox "Error " & err.Number & " (" & err.Description & ") in procedure AddObjectSelectionArea of Class Module Selection"
'
'End Sub

Sub Resize(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, Optional resizetype As Integer = 0, Optional Refresh As Boolean = True)
'first move selection to left and top, then resize
    parent.undoredo.ResetRedo
    Set undoch = New Changes

    Call Move(-boundaries.Left + Left, -boundaries.Top + Top, False)

    Call undoch.AddSelMove(-boundaries.Left + Left, -boundaries.Top + Top)

    Dim newselection(1023, 1023) As Integer
    Dim newbitfield As boolArray
    Set newbitfield = New boolArray
    
    Call newbitfield.InitArray(1024, 1024, 0)

    'get nrOfLinesOnX
    Dim nrOfLinesOnX As Integer
    nrOfLinesOnX = Round(((Bottom - Top) - (boundaries.Bottom - boundaries.Top)))
    Dim shrink As Boolean
    If nrOfLinesOnX < 0 Then
        shrink = True
        nrOfLinesOnX = Abs(nrOfLinesOnX)
    Else
        shrink = False
    End If

    Dim oldlinecount As Integer
    oldlinecount = 0

    Dim i As Integer
    Dim j As Integer
    For j = Top To Bottom
        If (j - Top) Mod Round((Bottom - Top) / nrOfLinesOnX) = 0 And j <> Top And j <> Bottom Then
            'we need to add/remove a line
            'check which of lines previous and next has the most
            '0 tiles
            Dim prevnr As Integer
            prevnr = TileCountOnLine(boundaries.Top + oldlinecount - 1, 0, False, boundaries)
            Dim nextnr As Integer
            nextnr = TileCountOnLine(boundaries.Top + oldlinecount, 0, False, boundaries)
            Dim curline As Integer
            Dim nextnextnr As Integer
            nextnextnr = TileCountOnLine(boundaries.Top + oldlinecount + 1, 0, False, boundaries)

            If shrink Then
                'skip the line
                oldlinecount = oldlinecount + 1

                'check which line to use
                If resizetype = 0 Then
                    curline = oldlinecount
                Else
                    If prevnr <= nextnr And prevnr <> 0 Then
                        curline = oldlinecount - 1
                    Else
                        curline = oldlinecount
                    End If
                End If

                'draw the line
                i = boundaries.Left
                For i = boundaries.Left To boundaries.Right
                    If isInSelection.value(i, boundaries.Top + curline) Then
                        If seltiles(i, boundaries.Top + curline) < 0 Then
                            newselection(i, j) = 0
                        Else
                            newselection(i, j) = seltiles(i, boundaries.Top + curline)
                        End If
                        newbitfield.value(i, j) = True
                    End If
                Next
                oldlinecount = oldlinecount + 1
            Else
                'check which line to use
                If resizetype = 0 Then
                    curline = oldlinecount
                ElseIf resizetype = 1 Then
                    ' check which line has the most tile 0
                    ' between prevnr and nextnr and use that line
                    If prevnr >= nextnr And prevnr <> 0 Then
                        curline = oldlinecount - 1
                    Else
                        curline = oldlinecount
                    End If
                ElseIf resizetype = 2 Then
                    If prevnr >= nextnr And prevnr <> 0 Then    'And boundaries.top + oldlinecount - 1 <= boundaries.bottom Then
                        If prevnr >= nextnextnr Then
                            curline = oldlinecount - 1
                        Else
                            curline = oldlinecount + 1
                        End If
                    Else
                        If nextnr >= nextnextnr Then
                            curline = oldlinecount
                        Else
                            curline = oldlinecount + 1
                        End If
                    End If
                End If

                'draw line
                i = boundaries.Left
                For i = boundaries.Left To boundaries.Right
                    If isInSelection.value(i, boundaries.Top + curline) Then
                        If seltiles(i, boundaries.Top + curline) < 0 Then
                            newselection(i, j) = 0
                        Else
                            newselection(i, j) = seltiles(i, boundaries.Top + curline)
                        End If
                        newbitfield.value(i, j) = True
                    End If
                Next
            End If
        Else
            i = boundaries.Left
            For i = boundaries.Left To boundaries.Right
                If isInSelection.value(i, boundaries.Top + oldlinecount) Then
                    If seltiles(i, boundaries.Top + oldlinecount) Then
                        newselection(i, j) = seltiles(i, boundaries.Top + oldlinecount)
                    Else
                        newselection(i, j) = seltiles(i, boundaries.Top + oldlinecount)
                    End If
                    newbitfield.value(i, j) = True
                End If
            Next
            oldlinecount = oldlinecount + 1
        End If
    Next


    Dim newnewselection(1023, 1023) As Integer
    
    Dim newnewbitfield As boolArray
    Set newnewbitfield = New boolArray
    
    Call newnewbitfield.InitArray(1024, 1024, 0)
    
'    Dim newnewbitfield(1023, 1023) As Boolean

    Dim nrOfLinesOnY As Integer
    nrOfLinesOnY = Round(((Right - Left) - (boundaries.Right - boundaries.Left)))
    If nrOfLinesOnY < 0 Then
        shrink = True
        nrOfLinesOnY = Abs(nrOfLinesOnY)
    Else
        shrink = False
    End If

    Dim bnds As area
    bnds.Left = Left
    bnds.Top = Top
    bnds.Right = boundaries.Right
    bnds.Bottom = Bottom

    oldlinecount = 0
    j = Left
    For j = Left To Right
        If (j - Left) Mod Round((Right - Left) / nrOfLinesOnY) = 0 And j <> Left And j <> Right Then
            'we need to add/remove a line
            'check which of lines previous and next has the most
            '0 tiles
            prevnr = TileCountOnLine(bnds.Left + oldlinecount - 1, 0, True, bnds)
            nextnr = TileCountOnLine(bnds.Left + oldlinecount, 0, True, bnds)
            nextnextnr = TileCountOnLine(bnds.Left + oldlinecount + 1, 0, True, bnds)
            If shrink Then
                'skip a line
                oldlinecount = oldlinecount + 1

                If prevnr <= nextnr And prevnr <> 0 Then    'And bnds.left + oldlinecount - 1 <= bnds.right Then
                    curline = oldlinecount - 1
                Else
                    curline = oldlinecount
                End If

                'draw the line
                i = bnds.Top
                For i = bnds.Top To bnds.Bottom
                    If newbitfield.value(bnds.Left + curline, i) Then
                        newnewselection(j, i) = newselection(bnds.Left + curline, i)
                        newnewbitfield.value(j, i) = True
                    End If
                Next
                oldlinecount = oldlinecount + 1
            Else
                'check which line to use
                If resizetype = 0 Then
                    curline = oldlinecount
                ElseIf resizetype = 1 Then
                    ' check which line has the most tile 0
                    ' between prevnr and nextnr and use that line
                    If prevnr >= nextnr And prevnr <> 0 Then
                        curline = oldlinecount - 1
                    Else
                        curline = oldlinecount
                    End If
                ElseIf resizetype = 2 Then
                    If prevnr >= nextnr And prevnr <> 0 Then
                        If prevnr >= nextnextnr Then
                            curline = oldlinecount - 1
                        Else
                            curline = oldlinecount + 1
                        End If
                    Else
                        If nextnr >= nextnextnr Then
                            curline = oldlinecount
                        Else
                            curline = oldlinecount + 1
                        End If
                    End If
                End If

                'draw the line
                i = bnds.Top
                For i = bnds.Top To bnds.Bottom
                    If newbitfield.value(bnds.Left + curline, i) Then
                        newnewselection(j, i) = newselection(bnds.Left + curline, i)
                        newnewbitfield.value(j, i) = True
                    End If
                Next
            End If
        Else
            i = bnds.Top
            For i = bnds.Top To bnds.Bottom
                If newbitfield.value(bnds.Left + oldlinecount, i) Then
                    newnewselection(j, i) = newselection(bnds.Left + oldlinecount, i)
                    newnewbitfield.value(j, i) = True
                End If
            Next
            oldlinecount = oldlinecount + 1
        End If
    Next

    bnds.Left = Left
    bnds.Top = Top
    bnds.Right = Right
    bnds.Bottom = Bottom

    Call DeleteSelection(undoch, False, True)



    seltiles = newnewselection
    Set isInSelection = newnewbitfield
    boundaries = bnds


    parent.pic1024selection.Cls
    Dim k As Integer
    Dim l As Integer
    For l = bnds.Top To bnds.Bottom
        For k = bnds.Left To bnds.Right
            Call SetPixel(parent.pic1024selection.hDC, k, l, TilePixelColor(seltiles(k, l)))

            If isInSelection.value(k, l) Then

                If seltiles(k, l) = 217 Or seltiles(k, l) = 219 Or seltiles(k, l) = 220 Then
                    Call undoch.AddTileChange(SelNew, k, l, seltiles(k, l))
                    Call CompleteSelObject(Me, k, l, undoch, False, True, True)
                ElseIf seltiles(k, l) >= 0 Then
                    Call undoch.AddTileChange(SelNew, k, l, seltiles(k, l))
                End If
            End If
        Next
    Next

    Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_RESIZE)

    hasAlreadySelectedParts = True

    parent.pic1024selection.Refresh

    selectionchanged = True

    If Refresh Then parent.UpdateLevel
End Sub

Friend Function TileCountOnLine(j As Integer, tile As Integer, vert As Boolean, boundaries As area)
    Dim i As Integer
    Dim count As Integer

    If vert Then
        For i = boundaries.Top To boundaries.Bottom
            If seltiles(j, i) = tile Then
                count = count + 1
            End If
        Next
    Else
        For i = boundaries.Left To boundaries.Right
            If seltiles(i, j) = tile Then
                count = count + 1
            End If
        Next
    End If
    TileCountOnLine = count
End Function

Sub ConverttoWalltiles(wallset As Integer, converttype As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim val As Integer
    Dim tmp As Integer
    Dim checked() As Boolean

    '    Dim isvalid As Boolean

  '  Erase checked
    ReDim checked(boundaries.Left - 1 To boundaries.Right, boundaries.Top - 1 To boundaries.Bottom)

    parent.undoredo.ResetRedo
    Set undoch = New Changes

    For j = boundaries.Top To boundaries.Bottom
        For i = boundaries.Left To boundaries.Right
            'Check if it is possible to convert current tile


            If IsValidForConvert(i, j, converttype, wallset) Then
                val = 0
                'Check right and under current tile
                If IsValidForConvert(i + 1, j, converttype, wallset) Then val = val + 1
                If IsValidForConvert(i, j + 1, converttype, wallset) Then val = val + 8

                'Check left and over if we already have checked this tile
                'This is to prevent problems when special tiles are used as walltiles (such as doors)
                If checked(i, j - 1) Then val = val + 2
                If checked(i - 1, j) Then val = val + 4

                'Set the tile if a walltile is defined for that value
                tmp = parent.walltiles.getWallTile(wallset, val)
                If tmp <> 0 Then seltiles(i, j) = tmp

                'Confirm that the tile was valid for conversion
                checked(i, j) = True
            End If
        Next
    Next

    parent.UpdateLevel

    Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_TOWALLTILE)

    selectionchanged = True

End Sub

Function IsValidForConvert(X As Integer, Y As Integer, convertmode As Integer, wallset As Integer) As Boolean
    Dim tilenr As Integer
    tilenr = seltiles(X, Y)
    If tilenr >= 0 And tilenr <> 217 And tilenr <> 219 And tilenr <> 220 Then
        If isInSelection.value(X, Y) And _
           (convertmode = 0 Or _
            (convertmode = 1 And parent.walltiles.GetTileIsWalltile(seltiles(X, Y)))) Or _
            (convertmode = 2 And parent.walltiles.GetTileIsWalltile(seltiles(X, Y), wallset)) Then

            IsValidForConvert = True
            Exit Function
        End If
    End If
    IsValidForConvert = False
End Function

Sub SelectAllTiles(inScreen As Boolean, add As Boolean)

    frmGeneral.IsBusy("Selection.SelectAllTiles") = True
    
    parent.undoredo.ResetRedo
    Set undoch = New Changes

    If Not add Then
        'deselect first
        Call ApplySelection(undoch, False)
    End If

    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y

    If inScreen Then
        lbx = parent.hScr.value \ parent.currenttilew
        hbx = (parent.hScr.value + parent.picpreview.width) \ parent.currenttilew
        lby = parent.vScr.value \ parent.currenttilew
        hby = (parent.vScr.value + parent.picpreview.height) \ parent.currenttilew
        
            
        If lbx < 0 Then lbx = 0
        If lby < 0 Then lby = 0
        If hbx > 1023 Then hbx = 1023
        If hby > 1023 Then hby = 1023
    Else
        lbx = 0
        lby = 0
        hbx = 1023
        hby = 1023
    End If

    
    Dim i As Integer
    Dim j As Integer

    For i = lbx To hbx
        For j = lby To hby
            If parent.getTile(i, j) <> 0 Then
                Call AddSelectionArea(i, i, j, j, undoch, False)
                selstate = Append
                hasAlreadySelectedParts = True
            End If
        Next
    Next

    Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTALLTILES)

    Call parent.UpdateLevel
    
    Call frmGeneral.UpdateToolBarButtons
    
    parent.tempdropping = False
    
    Call frmGeneral.SetCurrentTool(T_selection)
    
    frmGeneral.IsBusy("Selection.SelectAllTiles") = False
    
End Sub

Sub SelectTiles(tilenr As Integer, inScreen As Boolean)
    frmGeneral.IsBusy("Selection.SelectTiles") = True
    
    parent.undoredo.ResetRedo
    Set undoch = New Changes

    Call ApplySelection(undoch, False)

    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y

    If inScreen Then
        lbx = parent.hScr.value \ parent.currenttilew
        hbx = (parent.hScr.value + parent.picpreview.width) \ parent.currenttilew
        lby = parent.vScr.value \ parent.currenttilew
        hby = (parent.vScr.value + parent.picpreview.height) \ parent.currenttilew
        
        If lbx < 0 Then lbx = 0
        If lby < 0 Then lby = 0
        If hbx > 1023 Then hbx = 1023
        If hby > 1023 Then hby = 1023
    Else
        lbx = 0
        lby = 0
        hbx = 1023
        hby = 1023
    End If


    
    Dim i As Integer
    Dim j As Integer

    For i = lbx To hbx
        For j = lby To hby
            If parent.getTile(i, j) = tilenr Then
                Call AddSelectionArea(i, i, j, j, undoch, False)

            End If
        Next
    Next

    Call parent.UpdateLevel
    
    Call frmGeneral.UpdateToolBarButtons
    
    parent.tempdropping = False
    
    Call frmGeneral.SetCurrentTool(T_selection)
    
    frmGeneral.IsBusy("Selection.SelectTiles") = False

End Sub

Sub AddTiles(tilenr As Integer, inScreen As Boolean)
    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y
    Dim addedtiles As Boolean

    frmGeneral.IsBusy("Selection.AddTiles") = True
    
    parent.undoredo.ResetRedo
    Set undoch = New Changes

    If inScreen Then
        lbx = parent.hScr.value \ parent.currenttilew
        hbx = (parent.hScr.value + parent.picpreview.width) \ parent.currenttilew
        lby = parent.vScr.value \ parent.currenttilew
        hby = (parent.vScr.value + parent.picpreview.height) \ parent.currenttilew
            
        If lbx < 0 Then lbx = 0
        If lby < 0 Then lby = 0
        If hbx > 1023 Then hbx = 1023
        If hby > 1023 Then hby = 1023
    Else
        lbx = 0
        hbx = 1023
        lby = 0
        hby = 1023
    End If


    Dim i As Integer
    Dim j As Integer

    For i = lbx To hbx
        For j = lby To hby
            If parent.getTile(i, j) = tilenr Then
                Call AddTileToSelection(i, j, undoch)
                addedtiles = True
            End If
        Next
    Next

    If addedtiles Then
        Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTTILENR & " " & tilenr)

        Call parent.UpdateLevel
        
        Call frmGeneral.UpdateToolBarButtons
        
        parent.tempdropping = False
        
        Call frmGeneral.SetCurrentTool(T_selection)
        
    End If
    
    frmGeneral.IsBusy("Selection.AddTiles") = False

End Sub

Sub RemoveTiles(tilenr As Integer, inScreen As Boolean)
    'Removes all tiles <tilenr> from selection
    
    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y

    frmGeneral.IsBusy("Selection.RemoveTiles") = True
    
    parent.undoredo.ResetRedo
    Set undoch = New Changes

    If inScreen Then
        lbx = parent.hScr.value \ parent.currenttilew
        hbx = (parent.hScr.value + parent.picpreview.width) \ parent.currenttilew
        lby = parent.vScr.value \ parent.currenttilew
        hby = (parent.vScr.value + parent.picpreview.height) \ parent.currenttilew
            
        If lbx < 0 Then lbx = 0
        If lby < 0 Then lby = 0
        If hbx > 1023 Then hbx = 1023
        If hby > 1023 Then hby = 1023
    Else
        lbx = 0
        lby = 0
        hbx = 1023
        hby = 1023
    End If


    Dim i As Integer
    Dim j As Integer

    For i = lbx To hbx
        For j = lby To hby
            If parent.sel.getSelTile(i, j) = tilenr Then
                Call RemoveSelectionArea(i, i, j, j, undoch, False)
            End If
        Next
    Next

    Call parent.undoredo.AddToUndo(undoch, UNDO_UNSELECTTILENR & " " & tilenr)

    Call parent.UpdateLevel
    
    Call frmGeneral.UpdateToolBarButtons
    
    parent.tempdropping = False
    
    Call frmGeneral.SetCurrentTool(T_selection)
    
    frmGeneral.IsBusy("Selection.RemoveTiles") = False

End Sub



Sub DoCenterSelection(inScreen As Boolean)
    Dim undoch As New Changes
    
    Call StartMove
    
    Call CenterSelection(inScreen, True, undoch)
    
    Call EndMove(undoch)
    
    Call parent.undoredo.AddToUndo(undoch, UNDO_SELECTION_MOVE)

End Sub



Private Sub CenterSelection(inScreen As Boolean, Refresh As Boolean, undoch As Changes)
    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y

    If inScreen Then
        
        lbx = parent.hScr.value \ parent.currenttilew
        hbx = (parent.hScr.value + parent.picpreview.width) \ parent.currenttilew
        lby = parent.vScr.value \ parent.currenttilew
        hby = (parent.vScr.value + parent.picpreview.height) \ parent.currenttilew
            
        If lbx < 0 Then lbx = 0
        If lby < 0 Then lby = 0
        If hbx > 1023 Then hbx = 1023
        If hby > 1023 Then hby = 1023
    Else
        lbx = 0
        lby = 0
        hbx = 1023
        hby = 1023
    End If


    Dim csx As Integer
    Dim csy As Integer
    csx = lbx + (hbx - lbx) \ 2
    csy = lby + (hby - lby) \ 2

    Dim cx As Integer
    Dim cy As Integer
    cx = boundaries.Left + (boundaries.Right - boundaries.Left) \ 2
    cy = boundaries.Top + (boundaries.Bottom - boundaries.Top) \ 2
    
    Call StartMove
    
    Call Move(csx - cx, csy - cy, Refresh, True)
    
    Call EndMove(undoch)
    
            
    Call parent.UpdatePreview(Refresh)
End Sub


Sub DropTileFromSelection(X As Integer, Y As Integer, undoch As Changes)

    Dim oldtile As Integer
    oldtile = parent.getTile(X, Y)

    Call parent.setTile(X, Y, seltiles(X, Y), Nothing, False)
    Call parent.UpdateLevelTile(X, Y, False, False)

    Call undoch.AddTileChange(SelDrop, X, Y, oldtile)

    isInSelection.value(X, Y) = False
    Call setSelTile(X, Y, 0, Nothing, True, False)
End Sub

Sub AddTileToSelection(X As Integer, Y As Integer, undoch As Changes)
'adds map tile to selection

    Dim tmp As Integer

    tmp = parent.getTile(X, Y)

    Call setSelTile(X, Y, tmp, Nothing, True, False)

    If tmp <> 0 Then
        Call parent.setTile(X, Y, 0, Nothing, False)
        Call parent.UpdateLevelTile(X, Y, False, False)
    End If

    hasAlreadySelectedParts = True
    c_selstate = Append
    
    isInSelection.value(X, Y) = True

    Call undoch.AddTileChange(SelAdd, X, Y, 0)

End Sub

Sub CreateTileInSelection(X As Integer, Y As Integer, tilenr As Integer, undoch As Changes)
    If boundaries.Left > X Then boundaries.Left = X
    If boundaries.Right < X Then boundaries.Right = X
    If boundaries.Top > Y Then boundaries.Top = Y
    If boundaries.Bottom < Y Then boundaries.Bottom = Y

    If isInSelection.value(X, Y) Then
        Call setSelTile(X, Y, tilenr, undoch)
    Else
        Call setinselection(X, Y, True)
        Call setSelTile(X, Y, tilenr, Nothing, True, False)

        Call undoch.AddTileChange(SelNew, X, Y, tilenr)
    End If

    hasAlreadySelectedParts = True
End Sub

Sub DeleteSelectionTile(X As Integer, Y As Integer, undoch As Changes)

    If isInSelection.value(X, Y) Then
        Call setinselection(X, Y, False)
        Call undoch.AddTileChange(SelDelete, X, Y, seltiles(X, Y))
        Call setSelTile(X, Y, 0, Nothing, True, False)
    End If
End Sub

Sub setinselection(X As Integer, Y As Integer, val As Boolean)
    isInSelection.value(X, Y) = val
End Sub

Sub AddObjectToSelection(X As Integer, Y As Integer, undoch As Changes)
    Dim coords() As Integer
    Dim i As Integer
    Dim j As Integer
    Dim size As Integer

    coords = SearchObject(parent, X, Y)
    size = GetMaxSizeOfObject(parent.getTile(coords(0), coords(1)))

    For i = coords(0) To coords(0) + size
        For j = coords(1) To coords(1) + size
            Call AddTileToSelection(i, j, undoch)
        Next
    Next

    If boundaries.Left > coords(0) Then boundaries.Left = coords(0)
    If boundaries.Right < coords(0) + size Then boundaries.Right = coords(0) + size
    If boundaries.Top > coords(1) Then boundaries.Top = coords(1)
    If boundaries.Bottom < coords(1) + size Then boundaries.Bottom = coords(1) + size
End Sub


Sub DropObjectFromSelection(X As Integer, Y As Integer, undoch As Changes)
    Dim coords() As Integer
    Dim i As Integer
    Dim j As Integer
    Dim size As Integer

    coords = SearchSelObject(Me, X, Y)
    size = GetMaxSizeOfObject(seltiles(coords(0), coords(1)))

    For i = coords(0) To coords(0) + size
        For j = coords(1) To coords(1) + size
            If TileIsSpecial(parent.getTile(i, j)) Then
                Call SearchAndDestroyObject(parent, i, j, undoch, False)
            End If
        Next
    Next

    For i = coords(0) To coords(0) + size
        For j = coords(1) To coords(1) + size
            Call DropTileFromSelection(i, j, undoch)
        Next
    Next

End Sub


'Handles the undo/redo parts of moving the selection
Sub StartMove()
    moveboundaries = boundaries
End Sub

Sub EndMove(undoch As Changes)
    If Not undoch Is Nothing Then
        Call undoch.AddSelMove(boundaries.Left - moveboundaries.Left, boundaries.Top - moveboundaries.Top)
    End If
End Sub

Private Sub Class_Terminate()
    Erase seltiles
    Set parent = Nothing
    Set isInSelection = Nothing
    Set undoch = Nothing
End Sub
