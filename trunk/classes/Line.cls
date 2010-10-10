VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Dim startlinex As Integer
Dim startliney As Integer

Dim firstStartLineX As Integer, firstStartLineY As Integer


Dim undoch As Changes

Dim oldtilex As Integer
Dim oldtiley As Integer

Dim objfound As Boolean

Dim tooloptions As DrawOptions
Dim curSelection As TilesetSelection

Dim tagonly As Boolean


Private Sub Class_Initialize()
    oldtilex = -1
    oldtiley = -1
End Sub

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Friend Sub SetOptions(newoptions As DrawOptions)
    'Sets currently used tool options
    'This must be used before using DrawLine or other similar methods
    tooloptions = newoptions
End Sub

Friend Sub SetSelection(newselection As TilesetSelection)
    'Sets currently used tiles/object
    curSelection = newselection
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
    On Error GoTo MouseDown_Error
    
    If Button = vbLeftButton Or Button = vbRightButton Then
        Dim curtilex As Integer, curtiley As Integer
        curtilex = parent.ScreenToTileX(CInt(X))
        curtiley = parent.ScreenToTileY(CInt(Y))
        
        If curtilex < 0 Then curtilex = 0 Else If curtilex > 1023 Then curtilex = 1023
        If curtiley < 0 Then curtiley = 0 Else If curtiley > 1023 Then curtiley = 1023
    
        oldtilex = curtilex
        oldtiley = curtiley
    
        parent.undoredo.ResetRedo
        Set undoch = New Changes
    
        startlinex = curtilex
        startliney = curtiley
        
        firstStartLineX = startlinex
        firstStartLineY = startliney
        
        'get tileset selection
        curSelection = parent.tileset.selection(Button)
        
        ''''''''''''''''''
        'get tool options'
        If curtool = T_customshape Then
            toolOptions.step = 0
        Else
            toolOptions.step = frmGeneral.toolStep(curtool - 1).value
        End If
        
        'If curSelection.selectionType = TS_Tiles Or curSelection.selectionType = TS_Walltiles Then
        If curSelection.tileSize.X <> 1 Or curSelection.tileSize.Y <> 1 Or _
            (curSelection.selectionType <> TS_Tiles And curSelection.selectionType <> TS_Walltiles) Then
            toolOptions.size = 1
        Else
              If curtool = T_customshape Then
                  toolOptions.size = frmGeneral.customShapeSize(curCustomShape).value
              Else
                  toolOptions.size = frmGeneral.toolSize(curtool - 1).value
              End If
        End If
        
        
'''    If Not ignorethickness And Not multTileX > 1 And Not multTileY > 1 And tilenr <> 217 And tilenr <> 219 And tilenr <> 220 And Not parent.usingLVZTiles Then
'''        If curtool = T_line Or curtool = T_rectangle Or curtool = T_ellipse Or _
'''           curtool = t_spline Or curtool = T_pencil Or curtool = T_filledellipse Or _
'''           curtool = T_filledrectangle Or curtool = T_Eraser Then
'''            width = frmGeneral.toolSize(SharedVar.curtool - 1).Value
'''        ElseIf curtool = T_customshape Then
'''            If curCustomShape = s_cogwheel Or curCustomShape = s_star Or curCustomShape = s_regular Then
'''                width = frmGeneral.customShapeSize(curCustomShape).Value
'''            Else
'''                width = linew
'''            End If
'''        Else
'''            width = linew
'''        End If
'''    Else
'''        width = 1
'''    End If
'''    'TOREMOVE
'''    toolOptions.size = frmGeneral.toolSize(curtool - 1).Value
    
    
        toolOptions.filled = (curtool = T_filledellipse Or curtool = T_filledrectangle)
        
        If curtool = T_customshape Or curtool = T_filledellipse Or curtool = T_ellipse Then
            'ignore thickness while onpreview
            toolOptions.ignorethickness = frmGeneral.chkRenderAfter(curtool - 1).value = vbChecked
        Else
            'always show thickness
            toolOptions.ignorethickness = False
        End If
        
        If curtool = T_filledrectangle Or curtool = T_line Or curtool = T_pencil Or _
           curtool = T_rectangle Or curtool = t_spline Then
            If frmGeneral.optToolRound(curtool - 1).value Then
                toolOptions.drawshape = DS_Circle
            Else
                toolOptions.drawshape = DS_Rectangle
            End If
        Else
            toolOptions.drawshape = DS_Rectangle
        End If
        'we got tool options'
        '''''''''''''''''''''


    
        
        Call MouseMove(Button, X, Y, False)
    End If

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "Line.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single, usingctrl As Boolean)
    Dim curtilex As Integer
    Dim curtiley As Integer
    Dim fromtilex As Integer
    Dim fromtiley As Integer

    Static oldusectrl As Boolean

    On Error GoTo MouseMove_Error

    If oldtilex = -1 Or oldtiley = -1 Or SharedVar.MouseDown = 0 Then Exit Sub

    curtilex = parent.ScreenToTileX(CInt(X))
    curtiley = parent.ScreenToTileY(CInt(Y))

    If oldtilex = curtilex And oldtiley = curtiley And oldusectrl = usingctrl Then
        'nothing has moved
        Exit Sub
    End If


    If (curSelection.isSpecialObject) Then
        'ELLIPSE WITH SPECIAL OBJECTS??? forget it.
        If curtool = T_ellipse Or curtool = T_filledellipse Then
            Exit Sub
'        TOREMOVE
'        adjusting step shouldn't be necessary anymore...
'        Else
'            step = step + GetMaxSizeOfObject(tmptile)
        End If
    End If
    


    If usingctrl Then
        fromtilex = firstStartLineX
        fromtiley = firstStartLineY

        startlinex = fromtilex - (curtilex - fromtilex)
        startliney = fromtiley - (curtiley - fromtiley)
    Else
        startlinex = firstStartLineX
        startliney = firstStartLineY
    End If
    
    If curtilex < 0 Then curtilex = 0 Else If curtilex > 1023 Then curtilex = 1023
    If curtiley < 0 Then curtiley = 0 Else If curtiley > 1023 Then curtiley = 1023
    If startlinex < 0 Then startlinex = 0 Else If startlinex > 1023 Then startlinex = 1023
    If startliney < 0 Then startliney = 0 Else If startliney > 1023 Then startliney = 1023
    
    Call parent.UpdatePreview(False)
    
    Dim usewalltile As Boolean
    usewalltile = (curSelection.selectionType = TS_Walltiles)
    
    'Tag walltiles if needed
'    If usewalltile Then
'        Call parent.walltiles.ClearIsWallTile
'    End If

    If SharedVar.curtool = T_line Then
        Call DrawLine(startlinex, startliney, curtilex, curtiley, undoch, True, True, False)

    ElseIf SharedVar.curtool = T_rectangle Or SharedVar.curtool = T_filledrectangle Then
        Call DrawTileRectangle(startlinex, startliney, curtilex, curtiley, undoch, True, False)
        
    ElseIf SharedVar.curtool = T_ellipse Or SharedVar.curtool = T_filledellipse Then
        Call DrawEllipse(startlinex, startliney, curtilex, curtiley, undoch, True, False)
        
    ElseIf SharedVar.curtool = T_customshape Then
        Select Case curCustomShape
        Case s_cogwheel
            '                If usewalltile Then Call DrawCogWheel(tmptile, startlinex, startliney, curtilex, curtiley, 10, 0, 0.5, frmGeneral.optToolRound(SharedVar.curtool - 1).Value, undoch, tmpTileSizeX, tmpTileSizeY, True, , , True)
            '                Call DrawCogWheel(tmptile, startlinex, startliney, curtilex, curtiley, 10, 0, 0.5, frmGeneral.optToolRound(SharedVar.curtool - 1).Value, undoch, tmpTileSizeX, tmpTileSizeY, True, , , False)
            Call DrawCogWheel(startlinex, startliney, curtilex, curtiley, frmGeneral.customShapeTeethNumber(curCustomShape).value, frmGeneral.customShapeTeethSize(curCustomShape).value / 100, undoch, True, False, False)

        Case s_star
            
            Call DrawStar(startlinex, startliney, curtilex, curtiley, frmGeneral.customShapeTeethNumber(curCustomShape).value, frmGeneral.customShapeTeethSize(curCustomShape).value / 100, undoch, True, False, False)

        Case s_regular
            
            Call DrawRegularShape(startlinex, startliney, curtilex, curtiley, frmGeneral.customShapeTeethNumber(curCustomShape).value, undoch, True, False, False)

        End Select
    End If

    oldtilex = curtilex
    oldtiley = curtiley
    oldusectrl = usingctrl

    On Error GoTo 0
    Exit Sub

MouseMove_Error:
    HandleError Err, "Line.MouseMove"
End Sub

Sub MouseUp(Button As Integer, X As Single, Y As Single, Optional usingctrl As Boolean)
    Dim curtilex As Integer
    Dim curtiley As Integer
    Dim fromtilex As Integer
    Dim fromtiley As Integer

    
    If oldtilex = -1 Or oldtiley = -1 Then Exit Sub

    On Error GoTo MouseUp_Error

    curtilex = parent.ScreenToTileX(CInt(X))
    curtiley = parent.ScreenToTileY(CInt(Y))

    If (curSelection.tilenr = 217 Or curSelection.tilenr = 219 Or curSelection.tilenr = 220) Then
        If curtool = T_ellipse Or curtool = T_filledellipse Then
            Exit Sub
        End If
    End If



    If usingctrl Then
        fromtilex = firstStartLineX
        fromtiley = firstStartLineY

        startlinex = fromtilex - (curtilex - fromtilex)
        startliney = fromtiley - (curtiley - fromtiley)
    Else
        startlinex = firstStartLineX
        startliney = firstStartLineY
    End If


    If curtilex < 0 Then curtilex = 0 Else If curtilex > 1023 Then curtilex = 1023
    If curtiley < 0 Then curtiley = 0 Else If curtiley > 1023 Then curtiley = 1023
    If startlinex < 0 Then startlinex = 0 Else If startlinex > 1023 Then startlinex = 1023
    If startliney < 0 Then startliney = 0 Else If startliney > 1023 Then startliney = 1023



    frmGeneral.IsBusy("Line.MouseUp") = True

    Dim usewalltile As Boolean
    usewalltile = (curSelection.selectionType = TS_Walltiles)
    
    'Tag walltiles if needed
    If usewalltile Then
        Call parent.walltiles.ClearIsWallTile
    End If
    
    objfound = False
    
    If SharedVar.curtool = T_line Then
        Call DrawLine(startlinex, startliney, curtilex, curtiley, undoch, False, True, True)
    
    ElseIf SharedVar.curtool = T_rectangle Or SharedVar.curtool = T_filledrectangle Then
        Call DrawTileRectangle(startlinex, startliney, curtilex, curtiley, undoch, False, True)
    
    ElseIf SharedVar.curtool = T_ellipse Or SharedVar.curtool = T_filledellipse Then
        Call DrawEllipse(startlinex, startliney, curtilex, curtiley, undoch, False, True)
    
    ElseIf SharedVar.curtool = T_customshape Then
        Select Case curCustomShape
        Case s_cogwheel
            Call DrawCogWheel(startlinex, startliney, curtilex, curtiley, frmGeneral.customShapeTeethNumber(s_cogwheel).value, frmGeneral.customShapeTeethSize(s_cogwheel).value / 100, undoch, False, True, False)
        Case s_star
            Call DrawStar(startlinex, startliney, curtilex, curtiley, frmGeneral.customShapeTeethNumber(s_star).value, frmGeneral.customShapeTeethSize(s_star).value / 100, undoch, False, True, False)
        Case s_regular
            Call DrawRegularShape(startlinex, startliney, curtilex, curtiley, frmGeneral.customShapeTeethNumber(curCustomShape).value, undoch, False, True, False)

        End Select
    End If
    
    Call parent.UpdateLevel
    
'    If parent.magnifier.UsingPixels Then
'        parent.UpdateLevel
'    Else
'        'if using empty tile, or if a special object had to be deleted,
'        'refresh completly
'        If curSelection.tilenr = 0 Or curSelection.tilenr = 256 Or objfound Then
'            parent.UpdateLevel
'        Else
'            parent.UpdatePreview
'        End If
'    End If

    Call parent.undoredo.AddToUndo(undoch, ToolName(curtool))


    Call parent.walltiles.ClearIsWallTile
    
    frmGeneral.IsBusy("Line.MouseUp") = False
    
    On Error GoTo 0
    Exit Sub

MouseUp_Error:
    frmGeneral.IsBusy("Line.MouseUp") = False
    HandleError Err, "Line.MouseUp"
End Sub

'Sub DrawLine(tilenr As Integer, _
'             startx As Integer, startY As Integer, _
'             endX As Integer, endY As Integer, _
'             round As Boolean, _
'             undoch As Changes, _
' _
'    multTileX As Integer, multTileY As Integer, _
' _
'    Optional onpreview As Boolean = True, _
'             Optional Refresh As Boolean = True, _
'             Optional appendundo As Boolean = True, _
'             Optional ignorethickness As Boolean = False, _
'             Optional step As Integer = 0, _
'             Optional tagwalltiles As Boolean = False, _
'             Optional Left As Boolean = True)

Sub DrawLine(startx As Integer, startY As Integer, _
             endX As Integer, endY As Integer, _
             undoch As Changes, _
             onpreview As Boolean, _
             Refresh As Boolean, _
             appendundo As Boolean)

    Dim StepX As Integer, StepY As Integer

    'If tilenr = 217 Or tilenr = 219 Or tilenr = 220 Then
    StepX = tooloptions.Step + curSelection.tileSize.X
    StepY = tooloptions.Step + curSelection.tileSize.Y


    Dim tmpendx As Integer, tmpendy As Integer
    If StepX <> 1 Then
        tmpendx = startx + ((endX - startx) \ StepX) * StepX
    Else
        tmpendx = endX
    End If
    
    If StepY <> 1 Then
        tmpendy = startY + ((endY - startY) \ StepY) * StepY
    Else
        tmpendy = endY
    End If

    If Abs(tmpendx - startx) - 1 > Abs(tmpendy - startY) Then
        StepY = 1
    ElseIf Abs(tmpendy - startY) - 1 > Abs(tmpendx - startx) Then
        StepX = 1
    End If

    endX = startx + ((endX - startx) \ StepX) * StepX
    endY = startY + ((endY - startY) \ StepY) * StepY
    
 
    If curSelection.selectionType = TS_Walltiles Then
        parent.walltiles.ClearIsWallTile
        tagonly = True
        Call DrawThickLine(startx, startY, endX, endY, StepX, StepY, undoch, onpreview, appendundo)
    End If
    tagonly = False
    Call DrawThickLine(startx, startY, endX, endY, StepX, StepY, undoch, onpreview, appendundo)
    
    If Refresh Then
        If Not onpreview Then
            If curSelection.tilenr = 256 Or curSelection.tilenr = 0 Then
                parent.UpdateLevel
            Else
                parent.UpdatePreview
            End If
        Else
            parent.picPreview.Refresh
        End If
    End If
    
    On Error GoTo 0
    Exit Sub

DrawLine_Error:
    HandleError Err, "Line.DrawLine"
End Sub

'Private Sub DrawThickLine(tilenr As Integer, w As Integer, _
'                  sX As Integer, sY As Integer, _
'                  eX As Integer, eY As Integer, _
'                  StepX As Integer, _
'                  StepY As Integer, _
'                  undoch As Changes, _
'                  multTileX As Integer, _
'                  multTileY As Integer, _
'                  round As Boolean, _
'                  Optional onpreview As Boolean = True, Optional Refresh As Boolean = True, _
'                  Optional appendundo As Boolean = True, _
'                  Optional tagwalltiles As Boolean = False, _
'                  Optional Left As Boolean = True)
Private Sub DrawThickLine(ByVal startx As Integer, ByVal startY As Integer, _
                  ByVal endX As Integer, ByVal endY As Integer, _
                  StepX As Integer, StepY As Integer, _
                  undoch As Changes, _
                  onpreview As Boolean, _
                  appendundo As Boolean)
    
    Dim i As Integer, j As Integer

    Dim offset As Integer
    offset = (tooloptions.Size - 1) \ 2

    startx = startx - offset
    startY = startY - offset
    endX = endX - offset
    endY = endY - offset
    If startY > endY Then
        Swap startY, endY
        Swap startx, endX
    End If

    Dim lbx As Integer, lby As Integer
    
    lbx = parent.hScr.value \ parent.currenttilew
    lby = parent.vScr.value \ parent.currenttilew

    'the rectangle end overlaps the start, so be sure to fill it (so no gap occurs)
    Dim fillsrc As Boolean
    fillsrc = Abs(endX - startx) < tooloptions.Size Or Abs(endY - startY) < tooloptions.Size

    Dim dst As Single, halfw As Single
    Dim drawcurrent As Boolean
    
    halfw = tooloptions.Size / 2
    
    'draw the lines amount = (w + w) *2
    For j = 0 To tooloptions.Size - 1
        For i = 0 To (tooloptions.Size - 1)
            If tooloptions.drawshape = DS_Circle Then
                dst = dist(i, j, halfw - 0.5, halfw - 0.5)
                drawcurrent = (dst <= halfw)
            Else
                drawcurrent = True
            End If
            
            If drawcurrent Then
                If fillsrc Then
                    If onpreview Then
                        Call DrawTiles(startx + i, startY + j, lbx, lby)
                    Else
                        Call SetTiles(startx + i, startY + j, undoch, appendundo)
                    End If
                End If


                
                If StepX > 1 Or StepY > 1 Then
                    Call DrawBaseLine(startx + i, startY + j, endX + i, endY + j, undoch, onpreview, False, appendundo)
                ElseIf j = 0 Or i = 0 Or j = tooloptions.Size - 1 Or i = tooloptions.Size - 1 Then
                    Call DrawBaseLine(startx + i, startY + j, endX + i, endY + j, undoch, onpreview, False, appendundo)
                ElseIf tooloptions.drawshape = DS_Circle And halfw - dst <= 2 Then
                    Call DrawBaseLine(startx + i, startY + j, endX + i, endY + j, undoch, onpreview, False, appendundo)
                End If

            End If
        Next
    Next

End Sub

'Private Sub DrawBaseLine(tilenr As Integer, _
'                 startx As Integer, startY As Integer, _
'                 endX As Integer, endY As Integer, _
'                 StepX As Integer, _
'                 StepY As Integer, _
'                 undoch As Changes, _
'                 multTileX As Integer, _
'                 multTileY As Integer, _
'                 Optional onpreview As Boolean = True, Optional Refresh As Boolean = True, _
'                 Optional appendundo As Boolean = True, _
'                 Optional tagwalltiles As Boolean = False, _
'                 Optional Left As Boolean = True)
Private Sub DrawBaseLine(ByVal startx As Integer, ByVal startY As Integer, _
                 ByVal endX As Integer, ByVal endY As Integer, _
                 undoch As Changes, _
                 onpreview As Boolean, Refresh As Boolean, _
                 appendundo As Boolean)


    On Error GoTo DrawBaseLine_Error
    
    ' Bresenham's line algorithm for Microsoft Visual Basic 6.0
    ' Implementation by Robert Lee  July, 2002 Public Domain


    Dim InitialX As Integer, InitialY As Integer
    Dim FinalX As Integer, FinalY As Integer

    Dim i As Integer, j As Integer

    Dim val As Integer, tmp As Integer

    Dim StepX As Integer, StepY As Integer
    StepX = curSelection.tileSize.X + tooloptions.Step
    StepY = curSelection.tileSize.Y + tooloptions.Step
    
    InitialX = startx
    InitialY = startY

    FinalX = endX
    FinalY = endY

    Dim correctStepX As Integer, correctStepY As Integer

    Dim Steep As Boolean
    Dim DeltaX As Integer, DeltaY As Integer, Delta As Integer
    Dim coord As Integer

    Dim distX As Integer, distY As Integer
    distX = Abs(endX - startx)
    distY = Abs(endY - startY)

    If distY > distX Then
        StepX = 1
    ElseIf distX > distY Then
        StepY = 1
    End If
        
    DeltaX = distX / StepX
    DeltaY = distY / StepY
    
    If FinalX >= InitialX Then
        correctStepX = StepX
    Else
        correctStepX = -StepX
    End If
    If FinalY >= InitialY Then
        correctStepY = StepY
    Else
        correctStepY = -StepY
    End If
    
    
        

    
    If distX < distY Then
        Steep = True
        Swap InitialX, InitialY
        Swap DeltaX, DeltaY
        Swap correctStepX, correctStepY
    Else
        Steep = False
    End If

    'lbx & lby are calculated before the for, for optimization
    Dim lbx As Integer, lby As Integer
    lbx = parent.hScr.value \ parent.currenttilew
    lby = parent.vScr.value \ parent.currenttilew

    Delta = (DeltaY + DeltaY) - DeltaX

    For coord = 0 To DeltaX - 1
        If onpreview Then
            If Steep Then
                Call DrawTiles(InitialY, InitialX, lbx, lby)
            Else
                Call DrawTiles(InitialX, InitialY, lbx, lby)
            End If
        Else
            If Steep Then
                Call SetTiles(InitialY, InitialX, undoch, appendundo)
            Else
                Call SetTiles(InitialX, InitialY, undoch, appendundo)
            End If
        End If

        While Delta >= 0
            InitialY = InitialY + correctStepY
            Delta = Delta - (DeltaX + DeltaX)
            Delta = Delta
        Wend
        InitialX = InitialX + correctStepX
        Delta = Delta + (DeltaY + DeltaY)
    Next coord

    If onpreview Then
        Call DrawTiles(FinalX, FinalY, lbx, lby)
    Else
        Call SetTiles(FinalX, FinalY, undoch, appendundo)
    End If
    
    If Refresh Then
        If Not onpreview Then
            If curSelection.tilenr = 0 Or (curtool = T_spline And objfound) Then
                parent.UpdateLevel
            Else
                parent.UpdatePreview
            End If
        Else
            parent.picPreview.Refresh
        End If
    End If

    On Error GoTo 0
    Exit Sub

DrawBaseLine_Error:
    HandleError Err, "Line.DrawBaseLine"
End Sub

'Private Sub DrawTiles(ByRef tilenr As Integer, ByRef InitialX As Integer, ByRef InitialY As Integer, ByRef lbx As Integer, ByRef lby As Integer, _
'              ByRef undoch As Changes, _
'              ByRef multTileX As Integer, _
'              ByRef multTileY As Integer, _
'              Optional ByRef onpreview As Boolean = True, Optional ByRef Refresh As Boolean = True, _
'              Optional ByRef appendundo As Boolean = True, _
'              Optional ByRef tagwalltiles As Boolean = False, _
'              Optional ByRef Left As Boolean = True)
Private Sub SetTiles(ByRef InitialX As Integer, ByRef InitialY As Integer, _
              ByRef undoch As Changes, _
              ByVal appendundo As Boolean)

    Dim i As Integer, j As Integer, m As Integer, n As Integer

    Dim tilenr As Integer
    Dim maptile As Integer
    
    Dim offsetX As Integer
    Dim offsetY As Integer

    offsetX = (curSelection.tileSize.X - 1) \ 2
    offsetY = (curSelection.tileSize.Y - 1) \ 2
    
    
    If curSelection.selectionType = TS_LVZ Then
        Dim curX As Integer, curY As Integer
        
        curX = (InitialX - offsetX) * TILEW
        curY = (InitialY - offsetY) * TILEH
        
        Static lastX As Integer
        Static lastY As Integer
        
        If curX = lastX And curY = lastY And Not (lastX = 0 And lastY = 0) Then
            'skip add, we already added it last time!
            'this might be quicker than iterating over every map object to check if there is
            'already one
            
        Else
            'TODO: REPLACE THIS
            Dim mapobj As LVZMapObject
            
            mapobj.mode = parent.lvz.MapObjectDefaultMode
            mapobj.layer = parent.lvz.MapObjectDefaultLayer

            mapobj.imgidx = curSelection.tilenr
            mapobj.X = curX
            mapobj.Y = curY
            
            Call parent.lvz.AddMapObjectToLVZ(curSelection.group, mapobj)
            'Call parent.lvz.AddSelectedImageAsMapObjectToLVZ(Left, (InitialX - offsetX) * TileW, (InitialY - offsetY) * TileW, 0, False)
        
            lastX = curX
            lastY = curY
          
'                      If mapobj.layer >= lyr_AfterTiles Then
'                          Call parent.RedrawLvzOverArea(curX, curY, curX + parent.lvz.getImageWidth(curSelection.group, mapobj.imgidx), curY + parent.lvz.getImageHeight(curSelection.group, mapobj.imgidx), False)
'                      Else
'                          Call parent.RedrawLvzUnderArea(curX, curY, curX + parent.lvz.getImageWidth(curSelection.group, mapobj.imgidx), curY + parent.lvz.getImageHeight(curSelection.group, mapobj.imgidx), False)
'                      End If
        End If
    
    ElseIf curSelection.selectionType = TS_Tiles Or curSelection.selectionType = TS_Walltiles Then
        'TODO: GET RID OF THIS
        Dim useWall As Boolean
        useWall = (curSelection.selectionType = TS_Walltiles)
        If useWall Then parent.walltiles.curwall = curSelection.group
        ''''''
        
        Dim curtilex As Integer, curtiley As Integer
        Dim limitX As Integer, limitY As Integer
        
        If curSelection.isSpecialObject Then
            limitX = 0
            limitY = 0
        Else
            limitX = curSelection.tileSize.X - 1
            limitY = curSelection.tileSize.Y - 1
        End If

        
        For j = 0 To limitY
            
            curtiley = InitialY + j - offsetY
            
            For i = 0 To limitX
    
                curtilex = InitialX + i - offsetX
                
                If curtilex >= 0 And curtiley >= 0 And _
                   curtilex <= 1023 And curtiley <= 1023 Then
    
                    If useWall Then
                        'TODO: walltiles...
                        Call parent.walltiles.setIsWallTile(curtilex, curtiley)
                    End If
                    
                    maptile = parent.getTile(curtilex, curtiley)
                    
                    If Not tagonly Then
                        If parent.pastetype <> p_under Or maptile = 0 Then
                           
                            If useWall Then
                                If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) Then
                                    tilenr = parent.walltiles.DrawSelWallTile(curtilex, curtiley, undoch, True, False)
                                Else
                                    tilenr = parent.walltiles.DrawWallTile(curtilex, curtiley, True, undoch, False, True)
                                End If
                            Else
                                tilenr = curSelection.tilenr + i + 19 * j
                            End If
    
                            
                            'If a part of a special object is in the selection, draw it in the selection
                            If parent.sel.hasAlreadySelectedParts Then
                                
                                Dim considerselection As Boolean
                                If tilenr = TILE_WORMHOLE Or tilenr = TILE_LRG_ASTEROID Or tilenr = TILE_STATION Then
                                    considerselection = False
                                    For m = curtilex - GetMaxSizeOfObject(tilenr) \ 2 To curtilex + (0.5 + GetMaxSizeOfObject(tilenr)) \ 2
                                        For n = curtiley - GetMaxSizeOfObject(tilenr) \ 2 To curtiley + (0.5 + GetMaxSizeOfObject(tilenr)) \ 2
                                            If parent.sel.getIsInSelection(m, n) Then considerselection = True
                                        Next
                                    Next
                                End If
                                If considerselection Then
                                    Call parent.sel.setinselection(curtilex, curtiley, True)
                                End If
                            End If
    
                            If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) Then
                                'Add tile on selection
                                
    
                                If curSelection.isSpecialObject Then
                                    If setSelObject(parent.sel, tilenr, curtilex, curtiley, undoch, True) Then
                                        objfound = True
                                    End If
                                Else
                                    If TileIsSpecial(parent.sel.getSelTile(curtilex, curtiley)) Then
                                        objfound = True
                                        Call SearchAndDestroySelObject(parent.sel, curtilex, curtiley, undoch)
                                    End If
                                    
                                    Call parent.sel.setSelTile(curtilex, curtiley, tilenr, undoch)
                                    
                                    'TOREMOVE (that's pretty useless I think, pencil updates preview after)
                                    If curtool = T_pencil And Not parent.magnifier.UsingPixels Then
                                        'Call parent.DrawTile(tilenr, curtileX * parent.currenttilew - parent.Hscr.Value, curtileY * parent.currenttilew - parent.Vscr.Value)
                                    End If
                                End If
                            Else
                                'Add on map
                                If curSelection.isSpecialObject Then
                                        'check if it's the same object

                                    If tilenr = maptile Then
                                        'same object, don't do anything
                                    ElseIf setObject(parent, tilenr, curtilex, curtiley, undoch, False) Then
                                        objfound = True
                                        'Call SearchAndDestroyObject(parent, curtilex, curtiley, undoch, False)
                                    End If
    
    
                                Else
                                    If TileIsSpecial(maptile) Then
                                        objfound = True
                                        Call SearchAndDestroyObject(parent, curtilex, curtiley, undoch, False)
                                    End If
    
                                    Call parent.setTile(curtilex, curtiley, tilenr, undoch, appendundo)
                                    Call parent.UpdateLevelTile(curtilex, curtiley, False, True)
                                End If
                            End If
                            
                            
                        End If
                    End If
                End If
            Next
        Next
    End If
    
End Sub

Sub DrawTiles(ByRef InitialX As Integer, ByRef InitialY As Integer, ByRef lbx As Integer, ByRef lby As Integer)
              
    Dim i As Integer, j As Integer
    Dim tilenr As Integer
'    Dim tnr As Integer
'
'    If curSelection.tilenr = 0 Then tnr = 256 Else tnr = curSelection.tilenr

    Dim offsetX As Integer, offsetY As Integer

    offsetX = (curSelection.tileSize.X - 1) \ 2
    offsetY = (curSelection.tileSize.Y - 1) \ 2
    
    Dim drawX As Integer, drawY As Integer
    
    
    If curSelection.selectionType = TS_LVZ Then
        
         Dim srcDC As Long
         
         With parent.lvz.getImageDefinition(curSelection.group, curSelection.tilenr)
         
            srcDC = parent.lvz.pichDCzoom(.picboxIdx, parent.magnifier.zoom)
         
'         If parent.magnifier.zoom = 1 Then
'             Set srcPic = parent.picLVZItem(parent.lvz.getImageDefinition(lvzidx, imgidx).picboxidx)
'         ElseIf parent.magnifier.zoom = 2 Then
'             Set srcPic = parent.picLVZItemZoom2(parent.lvz.getImageDefinition(lvzidx, imgidx).picboxidx)
'         ElseIf parent.magnifier.zoom = 1 / 2 Then
'             Set srcPic = parent.picLVZItemZoom1_2(parent.lvz.getImageDefinition(lvzidx, imgidx).picboxidx)
'         ElseIf parent.magnifier.zoom = 1 / 4 Then
'             Set srcPic = parent.picLVZItemZoom1_4(parent.lvz.getImageDefinition(lvzidx, imgidx).picboxidx)
'         Else
'             Set srcPic = parent.picLVZItemZoomPix(parent.lvz.getImageDefinition(lvzidx, imgidx).picboxidx)
'         End If
'            If parent.currenttilew = TileW Then
'
'                Call BitBlt(parent.picPreview.hdc, InitialX * parent.currenttilew - parent.Hscr.value, InitialY * parent.currenttilew - parent.Vscr.value, parent.lvz.getSelImageWidth(Left), parent.lvz.getSelImageHeight(Left), parent.picLVZItem(parent.lvz.getSelImageIdx(Left)).hdc, 0, 0, vbSrcCopy)
'            Else
             Call TransparentBlt(parent.picPreview.hDc, (InitialX - offsetX) * parent.currenttilew - parent.hScr.value, (InitialY - offsetY) * parent.currenttilew - parent.vScr.value, .picWidth * (parent.currenttilew / TILEW), .picHeight * (parent.currenttilew / TILEW), srcDC, 0, 0, vbBlack)
'         End If
        End With
        
    ElseIf curSelection.selectionType = TS_Tiles Or curSelection.selectionType = TS_Walltiles Then
        
        'TODO: GET RID OF THIS
        Dim useWall As Boolean
        useWall = (curSelection.selectionType = TS_Walltiles)
        
        If useWall Then parent.walltiles.curwall = curSelection.group
        ''''''
        
        Dim UsingPixels As Boolean
        UsingPixels = parent.magnifier.UsingPixels
        
        
        Dim curtilex As Integer, curtiley As Integer
        Dim limitX As Integer, limitY As Integer
        
        If curSelection.isSpecialObject Then
            limitX = 0
            limitY = 0
        Else
            limitX = curSelection.tileSize.X - 1
            limitY = curSelection.tileSize.Y - 1
        End If
        
        curtiley = InitialY - offsetY
        
        drawY = parent.TileToScreenY(curtiley)
        
        For j = 0 To limitY
            

            curtilex = InitialX - offsetX

            drawX = parent.TileToScreenX(curtilex)
            
            
            For i = 0 To limitX
    

              
    
                If curtilex >= 0 And curtiley >= 0 And _
                   curtilex <= 1023 And curtiley <= 1023 Then
    
                    If useWall Then
                        'TODO: walltiles...
                        Call parent.walltiles.setIsWallTile(curtilex, curtiley)
                    End If
                    
                    If Not tagonly Then
                        If parent.pastetype <> p_under Or parent.getTile(curtilex, curtiley) = 0 Then
                            'It's useless to calculate the walltiles correctly at pixel level on preview
                            If useWall And Not parent.magnifier.UsingPixels Then
                                If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) Then
                                    tilenr = parent.walltiles.DrawSelWallTile(curtilex, curtiley, undoch, True, True)
                                Else
                                    tilenr = parent.walltiles.DrawWallTile(curtilex, curtiley, True, undoch, True, True)
                                End If
                            Else
                                tilenr = curSelection.tilenr + i + 19 * j
                            End If
                            
                          If UsingPixels Then
                                If tilenr = TILE_LRG_ASTEROID Then
                                    Call DrawRectangle(parent.picPreview.hDc, (curtilex - lbx), (curtiley - lby), (curtilex - lbx + 1), curtiley - lby + 1, TilePixelColor(tilenr))
                                ElseIf tilenr = TILE_STATION Then
                                    Call DrawRectangle(parent.picPreview.hDc, (curtilex - lbx - 2), (curtiley - lby - 2), (curtilex - lbx + 3), curtiley - lby + 3, TilePixelColor(tilenr))
                                ElseIf tilenr = TILE_WORMHOLE Then
                                    Call DrawRectangle(parent.picPreview.hDc, (curtilex - lbx - 2), (curtiley - lby - 2), (curtilex - lbx + 2), curtiley - lby + 2, TilePixelColor(tilenr))
                                Else
                                    Call SetPixel(parent.picPreview.hDc, (curtilex - lbx), (curtiley - lby), TilePixelColor(tilenr))
                                End If
                          Else
                              Call parent.TileRender.DrawTile(tilenr, False, parent.picPreview.hDc, drawX, drawY, False, True)
                          End If

                            
                        End If
                    End If
                    
                End If
                
                curtilex = curtilex + 1
                drawX = drawX + parent.currenttilew
            Next
            
            curtiley = curtiley + 1
            drawY = drawY + parent.currenttilew
        Next
    End If
'    parent.picPreview.Refresh
End Sub

Private Sub Swap(ByRef X As Integer, ByRef Y As Integer)
    Dim tmp As Integer
    tmp = X
    X = Y
    Y = tmp

End Sub

Private Sub DrawTileRectangle(ByVal startx As Integer, ByVal startY As Integer, ByVal endX As Integer, ByVal endY As Integer, ByRef undoch As Changes, ByVal onpreview As Boolean, ByVal appendundo As Boolean)
    Dim tick As Long
    tick = GetTickCount
    
'    Dim i As Integer
    Dim j As Integer
    '    Dim oldendx As Integer
    '    Dim oldendy As Integer

    On Error GoTo DrawTileRectangle_Error

    Dim correctStepX As Integer, correctStepY As Integer
                        
    correctStepX = curSelection.tileSize.X + tooloptions.Step
    correctStepY = curSelection.tileSize.Y + tooloptions.Step
            
    
    endX = startx + ((endX - startx) \ correctStepX) * correctStepX
    endY = startY + ((endY - startY) \ correctStepY) * correctStepY

    Call DrawLine(startx, startY, endX, startY, undoch, onpreview, False, appendundo)
    Call DrawLine(endX, startY, endX, endY, undoch, onpreview, False, appendundo)
    Call DrawLine(endX, endY, startx, endY, undoch, onpreview, False, appendundo)
    Call DrawLine(startx, endY, startx, startY, undoch, onpreview, False, appendundo)

    'Fill rectangle if needed
    If tooloptions.Filled And Not onpreview Then

        If startx <> endX And startY <> endY Then
            If startx > endX Then
                Swap startx, endX
            End If

            If startY > endY Then
                Swap startY, endY
            End If
            
            'To minimize the number of DrawLine calls, make the longest lines possible
            If endX - startx > endY - startY Then
                'Larger in X, make horizontal lines
                For j = startY + correctStepY To endY - correctStepY Step correctStepY
                    Call DrawLine(startx + correctStepX, j, endX - correctStepX, j, undoch, onpreview, False, appendundo)
                Next
            Else
                'Larger in Y, make vertical lines
                For j = startx + correctStepX To endX - correctStepX Step correctStepX
                    Call DrawLine(j, startY + correctStepY, j, endY - correctStepY, undoch, onpreview, False, appendundo)
                Next
            End If
        End If
    End If

    If Not onpreview Then
      Call parent.RedrawBuffer(True)
'310           parent.UpdatePreview
    Else
        parent.picPreview.Refresh
    End If
    
    frmGeneral.Label6.Caption = "rect:" & GetTickCount - tick
    On Error GoTo 0
    Exit Sub

DrawTileRectangle_Error:
    HandleError Err, "Line.DrawTileRectangle"
End Sub

Private Sub DrawEllipse(ByVal X As Integer, ByVal Y As Integer, ByVal xx As Integer, ByVal yy As Integer, ByRef undoch As Changes, ByVal onpreview As Boolean, ByVal appendundo As Boolean)
    Dim cX#, cY#, px#, py#, xp#, yp#
    Dim theta#, dtheta#, a#, b#
    Dim i As Integer, j As Integer


    On Error GoTo DrawEllipse_Error

  If X > xx Then Swap X, xx
  If Y > yy Then Swap Y, yy
     
    cX = X + (xx - X) / 2#
    cY = Y + (yy - Y) / 2#
    a = (xx - X) / 2#
    b = (yy - Y) / 2#

    dtheta = PI / 30#

    For theta = 0# To PI / 2 + dtheta Step dtheta
        xp = a * Cos(theta)
        yp = b * Sin(theta)

        'If xp < cx Then
        '    xp = xp + 0.01
        'Else
        '    xp = xp - 0.01
        'End If
        'If yp < cy Then
        '    yp = yp + 0.01
        'Else
        '    yp = yp - 0.01
        'End If

      If theta <> 0 Then
          Call DrawLine(CInt(cX + xp), CInt(cY + yp), CInt(cX + px), CInt(cY + py), undoch, onpreview, False, appendundo)
      
          Call DrawLine(CInt(cX + xp), CInt(cY - yp), CInt(cX + px), CInt(cY - py), undoch, onpreview, False, appendundo)
         
          Call DrawLine(CInt(cX - xp), CInt(cY + yp), CInt(cX - px), CInt(cY + py), undoch, onpreview, False, appendundo)
          
          Call DrawLine(CInt(cX - xp), CInt(cY - yp), CInt(cX - px), CInt(cY - py), undoch, onpreview, False, appendundo)
          

      End If

        px = xp
        py = yp


    Next theta

    If a <> 0 And b <> 0 Then
     If tooloptions.Filled Then
         If Not onpreview Then
         
         
'                    Dim iCy As Integer
'                    iCy = CInt((yy - Y) / 2) + Y
'
'                    Dim root As Double
'
'                      'Center
''                      Call DrawLine(CInt(cx - a), iCy, CInt(cx + a), iCy, undoch, onpreview, False, True)
'
'210                   For j = CInt(cy - b) + 1 To CInt(cy + b) - 1 Step curSelection.tileSize.Y + toolOptions.step
'
'                        root = Math.Sqr(a * a * (1 - ((j - cy) * (j - cy)) / (b * b)))
'                        Call DrawLine(CInt(cx - root + 1), j, CInt(cx + root - 1), j, undoch, onpreview, False, True)
'
''                        Call DrawLine(CInt(cx - root), iCy - j, CInt(cx + root), iCy - j, undoch, onpreview, False, True)
'
'

'270                   Next
                  For j = Y To yy Step curSelection.tileSize.Y
                    For i = X To xx Step curSelection.tileSize.X
                        If ((((i - cX) * (i - cX)) / (a * a)) + (((j - cY) * (j - cY)) / (b * b))) < 1 Then
                            Call DrawLine(i, j, i, j, undoch, onpreview, False, True)
                        End If

                    Next
                  Next

            End If
        End If
        
    End If

    If Not onpreview Then
        parent.UpdatePreview
    Else
        parent.picPreview.Refresh
    End If
    
    On Error GoTo 0
    Exit Sub

DrawEllipse_Error:
    HandleError Err, "Line.DrawEllipse"
End Sub

Sub DrawCogWheel(ByVal startx As Integer, ByVal startY As Integer, ByVal endX As Integer, ByVal endY As Integer, ByVal numsegments As Integer, ByVal sizeTeeth As Single, ByRef undoch As Changes, onpreview As Boolean, appendundo As Boolean, drawShapePreview As Boolean)
    Dim cX As Single, cY As Single, radius As Single
    
    Dim num_segments As Integer
    
    num_segments = IIf(numsegments * 2 <= 0, 1, numsegments * 2)

    If drawShapePreview Then
        frmGeneral.piccustomshapePreview(s_cogwheel).Cls
        cX = frmGeneral.piccustomshapePreview(s_cogwheel).ScaleWidth \ 2
        cY = cX
        radius = cX
    Else
        cX = startx ' Int(startx + (endX - startx) / 2)
        cY = startY ' Int(startY + (endY - startY) / 2)
        radius = Math.Sqr(CDbl(endX - startx) * CDbl(endX - startx) + CDbl(endY - startY) * CDbl(endY - startY)) '(endX - startx) / 2)
    End If

    
    Dim X As Single, Y As Single, oldX As Single, oldY As Single
    
    Dim theta As Double, startAngle As Double, dtheta As Double, dPart As Double
    

    If num_segments = 0 Then
        dPart = 1
    Else
        dPart = (2# * PI) / num_segments
    End If



    dtheta = (2# * PI) / (num_segments * 8)
      
    
      startAngle = Atn2(endX - startx, endY - startY)
    theta = startAngle + dPart / 2



    For theta = startAngle To startAngle + 2# * PI + dtheta Step dtheta

        Dim r As Single

        If CInt((theta - startAngle) / dPart) Mod 2 = 0 Then
            r = radius
        Else
            r = radius - sizeTeeth * radius
        End If
          
      If theta = startAngle Then
          'First segment
          oldX = cX + r * Cos(theta)
          oldY = cY + r * Sin(theta)
      End If
      
        X = cX + r * Cos(theta)
        Y = cY + r * Sin(theta)

        'Me.Line (oldx, oldy)-(cx + r * Cos(theta), cy + r * Sin(theta))
        If drawShapePreview Then
            frmGeneral.piccustomshapePreview(s_cogwheel).Line (Int(oldX), Int(oldY))-(Int(X), Int(Y)), vbBlack
        Else
            Call DrawLine(Int(oldX), Int(oldY), Int(X), Int(Y), undoch, onpreview, False, appendundo)
        End If

        oldX = X
        oldY = Y
        theta = (theta + dtheta)
    Next

    'Me.Line (cx, cy)-(cx + (radius - sizeTeeth * radius) * Cos(startAngle), cy + (radius - 0.2 * radius) * Sin(startAngle)), vbRed
    '    Call DrawLine(tilenr, CInt(oldx), CInt(oldy), cx + (radius - sizeTeeth * radius) * Cos(startAngle), cy + (radius - sizeTeeth * radius) * Sin(startAngle), round, undoch, multTileX, multTileY, onpreview, False, appendundo, False, 0, tagwalltiles)


    If Not onpreview Then
        '    Debug.Print "theta = " & theta
        parent.UpdatePreview
    Else
        parent.picPreview.Refresh
    End If


End Sub

Sub DrawStar(startx As Integer, startY As Integer, endX As Integer, endY As Integer, ByVal numsegments As Integer, ByVal sizeTeeth As Single, ByRef undoch As Changes, onpreview As Boolean, appendundo As Boolean, drawShapePreview As Boolean)
    Dim cX As Single
    Dim cY As Single
    Dim radius As Single
    Dim num_segments As Integer
    num_segments = IIf(numsegments * 2 <= 0, 1, numsegments * 2)


    If drawShapePreview Then
        frmGeneral.piccustomshapePreview(s_star).Cls
        cX = frmGeneral.piccustomshapePreview(s_star).ScaleWidth \ 2
        cY = cX
        radius = cX
    Else
        cX = startx ' Int(startx + (endX - startx) / 2)
        cY = startY ' Int(startY + (endY - startY) / 2)
        radius = Math.Sqr(CDbl(endX - startx) * CDbl(endX - startx) + CDbl(endY - startY) * CDbl(endY - startY)) '(endX - startx) / 2)
    End If

    Dim X As Single
    Dim Y As Single
    Dim theta As Single
    Dim dtheta As Single
    Dim seg As Integer

    Dim oldX As Single
    Dim oldY As Single

    dtheta = (2 * PI) / num_segments
      
      Dim startAngle As Double
    
      startAngle = Atn2(endX - startx, endY - startY)
    theta = startAngle

    oldX = cX + radius * Cos(theta)
    oldY = cY + radius * Sin(theta)

    For seg = 1 To num_segments
        Dim r As Single

        If seg Mod 2 = 0 Then
            r = radius
        Else
            r = radius - sizeTeeth * radius
        End If

        theta = (theta + dtheta)

        X = cX + r * Cos(theta)
        Y = cY + r * Sin(theta)

        'Me.Line (oldx, oldy)-(cx + r * Cos(theta), cy + r * Sin(theta))
        If drawShapePreview Then
            frmGeneral.piccustomshapePreview(s_star).Line (Int(oldX), Int(oldY))-(Int(X), Int(Y)), vbBlack
        Else
            Call DrawLine(Int(oldX), Int(oldY), Int(X), Int(Y), undoch, onpreview, False, appendundo)
        End If

        oldX = X
        oldY = Y

    Next seg

    'Me.Line (cx, cy)-(cx + (radius - sizeTeeth * radius) * Cos(startAngle), cy + (radius - 0.2 * radius) * Sin(startAngle)), vbRed
    '    Call DrawLine(tilenr, CInt(oldx), CInt(oldy), cx + (radius - sizeTeeth * radius) * Cos(startAngle), cy + (radius - sizeTeeth * radius) * Sin(startAngle), round, undoch, multTileX, multTileY, onpreview, False, appendundo, False, 0, tagwalltiles)


    If Not onpreview Then
        '    Debug.Print "theta = " & theta
        parent.UpdatePreview
    Else
        parent.picPreview.Refresh
    End If

    
End Sub

Sub DrawRegularShape(X As Integer, Y As Integer, xx As Integer, yy As Integer, ByVal numsegments As Integer, undoch As Changes, onpreview As Boolean, appendundo As Boolean, drawShapePreview As Boolean)

    Dim cX As Single
    Dim cY As Single
    Dim radius As Single

    If drawShapePreview Then
        frmGeneral.piccustomshapePreview(s_regular).Cls
        cX = frmGeneral.piccustomshapePreview(s_regular).ScaleWidth \ 2
        cY = cX
        radius = cX
    Else
        cX = X ' Int(startx + (endX - startx) / 2)
        cY = Y ' Int(startY + (endY - startY) / 2)
        radius = Math.Sqr(CDbl(xx - X) * CDbl(xx - X) + CDbl(yy - Y) * CDbl(yy - Y)) '(endX - startx) / 2)
    End If

    Dim theta As Single
    Dim dtheta As Single
    Dim seg As Integer

    Dim oldX As Single
    Dim oldY As Single
    Dim newX As Single
    Dim newY As Single

    dtheta = (2 * PI) / numsegments

    '  Me.CurrentX = cx + radius
    '  Me.CurrentY = cy
    Dim startAngle As Double
    
    startAngle = Atn2(xx - X, yy - Y)
    theta = startAngle

    oldX = cX + radius * Cos(theta)
    oldY = cY + radius * Sin(theta)

    '  Debug.Print oldx
    For seg = 1 To numsegments
        theta = CSng(theta + dtheta)
        '   Debug.Print (theta / (2 * PI)) * 100

        newX = cX + radius * CSng(Cos(theta))
        newY = cY + radius * CSng(Sin(theta))

        'Me.Line -(cx + radius * Cos(theta), cy + radius * Sin(theta))
        If drawShapePreview Then
            frmGeneral.piccustomshapePreview(s_regular).Line (Int(oldX), Int(oldY))-(Int(newX), Int(newY)), vbBlack
        Else
            Call DrawLine(Int(oldX), Int(oldY), Int(newX), Int(newY), undoch, onpreview, False, appendundo)
        End If

        oldX = newX
        oldY = newY
    Next seg
    ' Debug.Print X
    ' obj.Line (cx, cy)-(cx + (radius - 0.2 * radius) * Cos(startAngle), cy + (radius - 0.2 * radius) * Sin(startAngle)), vbRed

    If Not onpreview Then
        '    Debug.Print "theta = " & theta
        parent.UpdatePreview
    Else
        parent.picPreview.Refresh
    End If

    
End Sub

Sub SetObjFound(Optional value As Boolean = True)
    objfound = value
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
