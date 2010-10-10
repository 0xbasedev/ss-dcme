VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pencil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oldpenciltilex As Integer
Dim oldpenciltiley As Integer
Dim parent As frmMain

Dim undoch As Changes

Dim toolOptions As DrawOptions
Dim curSelection As TilesetSelection

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    On Error GoTo MouseDown_Error

    If Button <> vbLeftButton And Button <> vbRightButton Then Exit Sub
    
    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023

    parent.undoredo.ResetRedo
    Set undoch = New Changes

    oldpenciltilex = curtilex
    oldpenciltiley = curtiley

    'get tileset selection
    curSelection = parent.tileset.selection(Button)
    If curtool = T_Eraser Then
        curSelection.selectionType = TS_Tiles
        curSelection.isSpecialObject = False
        curSelection.tilenr = 0
        curSelection.tileSize.X = 1
        curSelection.tileSize.Y = 1
        curSelection.pixelSize.X = TILEW
        curSelection.pixelSize.Y = TILEH
    End If
    
    Call parent.tline.SetSelection(curSelection)
    
    If frmGeneral.optToolRound(curtool - 1).value Then
        toolOptions.drawshape = DS_Circle
    Else
        toolOptions.drawshape = DS_Rectangle
    End If
    
    If curSelection.tileSize.X <> 1 Or curSelection.tileSize.Y <> 1 Then
        toolOptions.size = 1
    Else
        toolOptions.size = frmGeneral.toolSize(curtool - 1).value
    End If
        
    
    Call parent.tline.SetOptions(toolOptions)
    

'    'Advanced pencil mode
'    If (tmptile = 217 Or tmptile = 219 Or tmptile = 220) Then
'        If CBool(GetSetting("AdvancedPencil", "0")) Then
'            Call parent.setTile(curtilex - GetMaxSizeOfObject(tmptile) \ 2, curtiley - GetMaxSizeOfObject(tmptile) \ 2, tmptile, undoch, True)
'            Call parent.UpdateLevelTile(curtilex - GetMaxSizeOfObject(tmptile) \ 2, curtiley - GetMaxSizeOfObject(tmptile) \ 2)
'            Exit Sub
'        End If
'    End If
    Call MouseMove(Button, X, Y)
'    If tmptileSizeX > 1 Or tmptileSizeY > 1 Or tmptile = 217 Or tmptile = 219 Or tmptile = 220 Or parent.usingLVZTiles Then
'        Call MouseMove(button, x, y)
'    Else
'        If frmGeneral.toolSize(curtool - 1).Value <> 1 Then
'            Call MouseMove(button, x, y)
'        Else
'            If parent.pastetype <> p_under Or _
'               parent.pastetype = p_under And parent.getTile(curtilex, curtiley) = 0 Then
'
'                If parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley) Then
'                    If isSelTileSpecial(parent.sel, curtilex, curtiley) Then
'                        Call SearchAndDestroySelObject(parent.sel, curtilex, curtiley, undoch)
'                    End If
'
'                    If usewalltile Then
'                        Call parent.walltiles.ClearIsWallTile
'                        Call parent.walltiles.setIsWallTile(curtilex, curtiley)
'                        tmptile = parent.walltiles.DrawSelWallTile(curtilex, curtiley, undoch, True, False)
'                    End If
'                    Call parent.sel.setSelTile(curtilex, curtiley, tmptile, undoch)
'                    If Not parent.magnifier.UsingPixels Then Call parent.DrawTile(tmptile, (curtilex * parent.currenttilew - parent.Hscr.Value), (curtiley * parent.currenttilew - parent.Vscr.Value))
'                Else
'                    If isTileSpecial(parent, curtilex, curtiley) Then
'                        Call SearchAndDestroyObject(parent, curtilex, curtiley, undoch)
'                    End If
'
'                    If usewalltile Then
'                        Call parent.walltiles.ClearIsWallTile
'                        Call parent.walltiles.setIsWallTile(curtilex, curtiley)
'                        tmptile = parent.walltiles.DrawWallTile(curtilex, curtiley, True, undoch, False)
'                    End If
'                    Call parent.setTile(curtilex, curtiley, tmptile, undoch)
'                    Call parent.UpdateLevelTile(curtilex, curtiley, False)
'                End If
'            End If
'        End If
'    End If


'360       If parent.magnifier.UsingPixels Then
'370           Call parent.UpdateLevel
'380       Else
'390           Call parent.UpdatePreview
'400       End If
    

    Call parent.walltiles.ClearIsWallTile

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "Pencil.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer

    On Error GoTo MouseMove_Error

    If oldpenciltilex = -1 Or oldpenciltiley = -1 Then Exit Sub

    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023
    
        
'    If curtool = T_Eraser Then
'        tmptile = 0
'        tmptileSizeX = 1
'        tmptileSizeY = 1
'        usewalltile = False
'    Else
'
'        If parent.usingLVZTiles Then
'            Dim lvzwidth As Integer
'            Dim lvzheight As Integer
'            tmptile = 1
'            lvzwidth = parent.lvz.getSelImageWidth(button = vbLeftButton)
'            lvzheight = parent.lvz.getSelImageHeight(button = vbLeftButton)
'
'            tmptileSizeX = IIf(lvzwidth Mod TileW = 0, lvzwidth \ TileW, lvzwidth \ TileW + 1)
'            tmptileSizeY = IIf(lvzheight Mod TileW = 0, lvzheight \ TileW, lvzheight \ TileW + 1)
'            usewalltile = False
'
'        ElseIf button = vbLeftButton Then
'            tmptile = parent.tilesetleft
'            tmptileSizeX = parent.multTileLeftx
'            tmptileSizeY = parent.multTileLefty
'            usewalltile = parent.useWallTileLeft
'            parent.walltiles.curwall = parent.curWallTileLeft
'        ElseIf button = vbRightButton Then
'            tmptile = parent.tilesetright
'            tmptileSizeX = parent.multTileRightx
'            tmptileSizeY = parent.multTileRighty
'            usewalltile = parent.useWallTileRight
'            parent.walltiles.curwall = parent.curWallTileRight
'        End If
'    End If
    If Button Then
        If curSelection.selectionType = TS_Walltiles Then
            Call parent.walltiles.ClearIsWallTile
        End If

        
        'If using advanced pencil to do weird stuff
        If frmGeneral.chkAdvancedPencil.value = vbChecked And curSelection.selectionType = TS_Tiles Then
            If curSelection.isSpecialObject Then
                Call parent.setTile(curtilex - (curSelection.tileSize.X - 1) \ 2, curtiley - (curSelection.tileSize.Y - 1) \ 2, curSelection.tilenr, undoch, True)
                Call parent.UpdateLevelTile(curtilex - (curSelection.tileSize.X - 1) \ 2, curtiley - (curSelection.tileSize.Y - 1) \ 2)
                
            Else
                Call parent.setTile(curtilex, curtiley, curSelection.tilenr, undoch, True)
                Call parent.UpdateLevelTile(curtilex, curtiley)
            End If
            Exit Sub
        End If

        'draw a line from the old x,y s
        'its more fluent
        If oldpenciltilex >= 0 And oldpenciltiley >= 0 Then
            
            Call parent.tline.DrawLine(oldpenciltilex, oldpenciltiley, curtilex, curtiley, undoch, False, curSelection.selectionType = TS_LVZ, True)
        End If
        '        End If

      If curSelection.selectionType = TS_LVZ Then
          Call parent.RedrawLvzUnderLayer(False)
          Call parent.RedrawLvzOverLayer(False)
          Call parent.RedrawBuffer(False)
          Call parent.UpdatePreview(True, False)
      ElseIf parent.magnifier.UsingPixels Then
          Call parent.RedrawTileLayer(True)
'                Call parent.RedrawBuffer(False)
'                Call parent.UpdatePreview(True, False)
      Else
          Call parent.RedrawBuffer(False)
          Call parent.UpdatePreview(True, False)
      End If
      
    End If


    oldpenciltilex = curtilex
    oldpenciltiley = curtiley

    On Error GoTo 0
    Exit Sub

MouseMove_Error:
    HandleError Err, "Pencil.MouseMove"
End Sub

Sub MouseUp()
'do nothing
    On Error GoTo MouseUp_Error

    If oldpenciltilex = -1 Or oldpenciltiley = -1 Then Exit Sub

    Call parent.undoredo.AddToUndo(undoch, IIf(curtool = T_pencil, UNDO_PENCIL, UNDO_ERASER))
    'for tile 0 being gridded again we need to update the level
    Call parent.UpdateLevel

    On Error GoTo 0
    Exit Sub

MouseUp_Error:
    HandleError Err, "Pencil.MouseUp"
End Sub


Private Sub Class_Initialize()
    oldpenciltilex = -1
    oldpenciltiley = -1
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
