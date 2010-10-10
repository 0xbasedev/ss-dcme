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
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        On Error GoTo MouseDown_Error

20        If Button <> vbLeftButton And Button <> vbRightButton Then Exit Sub
          
30        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
40        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

50        If curtilex < 0 Then curtilex = 0
60        If curtiley < 0 Then curtiley = 0
70        If curtilex > 1023 Then curtilex = 1023
80        If curtiley > 1023 Then curtiley = 1023

90        parent.undoredo.ResetRedo
100       Set undoch = New Changes

110       oldpenciltilex = curtilex
120       oldpenciltiley = curtiley

          'get tileset selection
130       curSelection = parent.tileset.selection(Button)
140       If curtool = T_Eraser Then
150           curSelection.selectionType = TS_Tiles
160           curSelection.isSpecialObject = False
170           curSelection.tilenr = 0
180           curSelection.tileSize.X = 1
190           curSelection.tileSize.Y = 1
200           curSelection.pixelSize.X = TILEW
210           curSelection.pixelSize.Y = TILEH
220       End If
          
230       Call parent.tline.SetSelection(curSelection)
          
240       If frmGeneral.optToolRound(curtool - 1).value Then
250           toolOptions.drawshape = DS_Circle
260       Else
270           toolOptions.drawshape = DS_Rectangle
280       End If
          
290       If curSelection.tileSize.X <> 1 Or curSelection.tileSize.Y <> 1 Then
300           toolOptions.size = 1
310       Else
320           toolOptions.size = frmGeneral.toolSize(curtool - 1).value
330       End If
              
          
340       Call parent.tline.SetOptions(toolOptions)
          

      '    'Advanced pencil mode
      '    If (tmptile = 217 Or tmptile = 219 Or tmptile = 220) Then
      '        If CBool(GetSetting("AdvancedPencil", "0")) Then
      '            Call parent.setTile(curtilex - GetMaxSizeOfObject(tmptile) \ 2, curtiley - GetMaxSizeOfObject(tmptile) \ 2, tmptile, undoch, True)
      '            Call parent.UpdateLevelTile(curtilex - GetMaxSizeOfObject(tmptile) \ 2, curtiley - GetMaxSizeOfObject(tmptile) \ 2)
      '            Exit Sub
      '        End If
      '    End If
350       Call MouseMove(Button, X, Y)
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
          

410       Call parent.walltiles.ClearIsWallTile

420       On Error GoTo 0
430       Exit Sub

MouseDown_Error:
440       HandleError Err, "Pencil.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer

10        On Error GoTo MouseMove_Error

20        If oldpenciltilex = -1 Or oldpenciltiley = -1 Then Exit Sub

30        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
40        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

50        If curtilex < 0 Then curtilex = 0
60        If curtiley < 0 Then curtiley = 0
70        If curtilex > 1023 Then curtilex = 1023
80        If curtiley > 1023 Then curtiley = 1023
          
              
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
90        If Button Then
100           If curSelection.selectionType = TS_Walltiles Then
110               Call parent.walltiles.ClearIsWallTile
120           End If

              
              'If using advanced pencil to do weird stuff
130           If frmGeneral.chkAdvancedPencil.value = vbChecked And curSelection.selectionType = TS_Tiles Then
140               If curSelection.isSpecialObject Then
150                   Call parent.setTile(curtilex - (curSelection.tileSize.X - 1) \ 2, curtiley - (curSelection.tileSize.Y - 1) \ 2, curSelection.tilenr, undoch, True)
160                   Call parent.UpdateLevelTile(curtilex - (curSelection.tileSize.X - 1) \ 2, curtiley - (curSelection.tileSize.Y - 1) \ 2)
                      
170               Else
180                   Call parent.setTile(curtilex, curtiley, curSelection.tilenr, undoch, True)
190                   Call parent.UpdateLevelTile(curtilex, curtiley)
200               End If
210               Exit Sub
220           End If

              'draw a line from the old x,y s
              'its more fluent
230           If oldpenciltilex >= 0 And oldpenciltiley >= 0 Then
                  
240               Call parent.tline.DrawLine(oldpenciltilex, oldpenciltiley, curtilex, curtiley, undoch, False, curSelection.selectionType = TS_LVZ, True)
250           End If
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


340       oldpenciltilex = curtilex
350       oldpenciltiley = curtiley

360       On Error GoTo 0
370       Exit Sub

MouseMove_Error:
380       HandleError Err, "Pencil.MouseMove"
End Sub

Sub MouseUp()
      'do nothing
10        On Error GoTo MouseUp_Error

20        If oldpenciltilex = -1 Or oldpenciltiley = -1 Then Exit Sub

30        Call parent.undoredo.AddToUndo(undoch, IIf(curtool = T_pencil, UNDO_PENCIL, UNDO_ERASER))
          'for tile 0 being gridded again we need to update the level
40        Call parent.UpdateLevel

50        On Error GoTo 0
60        Exit Sub

MouseUp_Error:
70        HandleError Err, "Pencil.MouseUp"
End Sub


Private Sub Class_Initialize()
10        oldpenciltilex = -1
20        oldpenciltiley = -1
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
20        Set undoch = Nothing
End Sub
