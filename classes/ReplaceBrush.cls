VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplaceBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain
Dim undoch As Changes

Dim oldposX As Integer
Dim oldposY As Integer

Dim sourceButton As Integer  'Items to replace
Dim replaceButton As Integer 'Items to replace with

Dim srcSelection As TilesetSelection
Dim newselection As TilesetSelection

Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
      '    Dim curtilex As Integer
      '    Dim curtiley As Integer
      '
      '    curtilex = (parent.hScr.Value + X) \ parent.currenttilew
      '    curtiley = (parent.vScr.Value + Y) \ parent.currenttilew
      '
      '    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
      '        Exit Sub
      '    End If
10        If Button = vbLeftButton Then
20            sourceButton = vbLeftButton
30            replaceButton = vbRightButton
40        ElseIf Button = vbRightButton Then
50            sourceButton = vbRightButton
60            replaceButton = vbLeftButton
70        Else
80            Exit Sub
90        End If
          
100       oldposX = -1
110       oldposY = -1
          
120       srcSelection = parent.tileset.selection(sourceButton)
130       newselection = parent.tileset.selection(replaceButton)
          
140       parent.undoredo.ResetRedo
150       Set undoch = New Changes

160       Call MouseMove(Button, X, Y)
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer

10        If Button <> vbLeftButton And Button <> vbRightButton Then
20            Exit Sub
30        End If

40        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
50        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

60        If curtilex < 0 Then curtilex = 0
70        If curtiley < 0 Then curtiley = 0
80        If curtilex > 1023 Then curtilex = 1023
90        If curtiley > 1023 Then curtiley = 1023

100       If oldposX = -1 Or oldposY = -1 Then
110           oldposX = curtilex
120           oldposY = curtiley
130       ElseIf oldposX = curtilex And oldposY = curtiley Then
140           Exit Sub
150       End If

      '    If parent.multTileLeftx <> parent.multTileRightx Or _
      '       parent.multTileLefty <> parent.multTileRighty Then
      '        Exit Sub
      '    End If

      '    If parent.tilesetleft = 217 Or parent.tilesetleft = 219 Or _
      '       parent.tilesetleft = 220 Or _
      '       parent.tilesetright = 217 Or parent.tilesetright = 219 Or _
      '       parent.tilesetright = 220 Then
      '        Exit Sub
      '    End If
          
160       Call DrawReplaceLine(srcSelection.tilenr, newselection.tilenr, oldposX, oldposY, curtilex, curtiley, undoch, CInt(srcSelection.tileSize.X), CInt(srcSelection.tileSize.Y), frmGeneral.toolSize(curtool - 1).value, True, True)

170       oldposX = curtilex
180       oldposY = curtiley
End Sub

Sub MouseUp()
      'do nothing
10        If oldposX = -1 Or oldposY = -1 Then Exit Sub

20        Call parent.undoredo.AddToUndo(undoch, UNDO_REPLACEBRUSH)
          'for tile 0 being gridded again we need to update the level
30        Call parent.UpdateLevel

End Sub

Private Sub DrawReplaceLine(srctile As Integer, newtile As Integer, sX As Integer, sY As Integer, eX As Integer, eY As Integer, undoch As Changes, multTileX As Integer, multTileY As Integer, width As Integer, Optional Refresh As Boolean = True, Optional appendundo As Boolean = True, Optional tagwalltiles As Boolean = False)
          Dim i As Integer
          Dim j As Integer

          '    If tagwalltiles = True Then
          '        'We're using walltiles, define which tiles will be in the area to tile
          '        If width <> 1 And Not ignoreThickness And Not multTileX > 1 And Not multTileY > 1 Then
          '            For j = -(width \ 2) + 1 To (width \ 2)
          '                For i = -(width \ 2) + 1 To (width \ 2)
          '                    Call DrawBaseLine(1, startx + i, starty + j, endx + i, endy + j, stepX, stepY, undoch, multTileX, multTileY, onpreview, False, appendundo, tagwalltiles = True)
          '                Next
          '            Next
          '        Else
          '            Call DrawBaseLine(1, startx, starty, endx, endy, stepX, stepY, undoch, multTileX, multTileY, onpreview, refresh, appendundo, True)
          '        End If
          '
          '    Else
          Dim startx As Integer
          Dim startY As Integer
          Dim endX As Integer
          Dim endY As Integer
          Dim offset As Integer
10        offset = (width - 1) \ 2

20        startx = sX - offset
30        startY = sY - offset
40        endX = eX - offset
50        endY = eY - offset

60        If width <> 1 Then
70            For j = 0 To width - 1
80                For i = 0 To width - 1
90                    Call DrawBaseLine(srctile, newtile, startx + i, startY + j, endX + i, endY + j, undoch, multTileX, multTileY, False, appendundo)
100               Next
110           Next
120           If Refresh Then
130               If parent.magnifier.UsingPixels Then
140                   parent.UpdateLevel
150               Else
160                   parent.UpdatePreview
170               End If
180           End If
190       Else
200           Call DrawBaseLine(srctile, newtile, startx, startY, endX, endY, undoch, multTileX, multTileY, Refresh, appendundo)
210       End If

          '    End If
End Sub

Private Sub DrawBaseLine(srctile As Integer, newtile As Integer, startx As Integer, startY As Integer, endX As Integer, endY As Integer, undoch As Changes, multTileX As Integer, multTileY As Integer, Optional Refresh As Boolean = True, Optional appendundo As Boolean = True, Optional tagwalltiles As Boolean = False)
      ' Bresenham's line algorithm for Microsoft Visual Basic 6.0
      ' Implementation by Robert Lee <rlee0001@maine.rr.com> July, 2002 Public Domain
          Dim InitialX As Integer
          Dim InitialY As Integer
          Dim FinalX As Integer
          Dim FinalY As Integer

10        InitialX = startx
20        InitialY = startY
30        FinalX = endX
40        FinalY = endY

          Dim Steep As Boolean
          Dim DeltaX As Integer, DeltaY As Integer, Delta As Integer
          Dim StepX As Integer, StepY As Integer
          Dim coord As Integer

50        Steep = False
60        DeltaX = Abs(FinalX - InitialX)
70        If (FinalX - InitialX) > 0 Then
80            StepX = 1
90        Else
100           StepX = -1
110       End If
120       DeltaY = Abs(FinalY - InitialY)
130       If (FinalY - InitialY) > 0 Then
140           StepY = 1
150       Else
160           StepY = -1
170       End If
180       If DeltaY > DeltaX Then
190           Steep = True
200           Swap InitialX, InitialY
210           Swap DeltaX, DeltaY
220           Swap StepX, StepY
230       End If
240       Delta = (DeltaY * 2) - DeltaX
250       For coord = 0 To DeltaX - 1
260           If Steep Then
270               Call DrawTiles(srctile, newtile, InitialY, InitialX, undoch, multTileX, multTileY, appendundo, tagwalltiles)
280           Else
290               Call DrawTiles(srctile, newtile, InitialX, InitialY, undoch, multTileX, multTileY, appendundo, tagwalltiles)
300           End If
310           While Delta >= 0
320               InitialY = InitialY + StepY
330               Delta = Delta - (DeltaX * 2)
340           Wend
350           InitialX = InitialX + StepX
360           Delta = Delta + (DeltaY * 2)
370       Next coord

380       Call DrawTiles(srctile, newtile, FinalX, FinalY, undoch, multTileX, multTileY, appendundo, tagwalltiles)

390       If Refresh Then
400           parent.UpdatePreview
410       End If
End Sub

Private Sub DrawTiles(srctile As Integer, newtile As Integer, X As Integer, Y As Integer, undoch As Changes, multX As Integer, multY As Integer, Optional appendundo As Boolean = True, Optional tagwalltiles As Boolean = False)
          Dim i As Integer
          Dim j As Integer


          Dim curI As Integer
          Dim curJ As Integer
          Dim chI As Integer
          Dim chJ As Integer
          
10        For j = 0 To multY - 1
20            curJ = (srctile \ 19) + j
30            chJ = (newtile \ 19) + j
              
40            For i = 0 To multX - 1

50                curI = (srctile Mod 19) + i
60                chI = (newtile Mod 19) + i


70                If Not (X < 0 Or X > 1023 Or Y < 0 Or Y > 1023) Then
80                    If parent.sel.getIsInSelection(X, Y) Then
90                        If parent.sel.getSelTile(X, Y) = chJ * 19 + chI Then
100                           Call parent.sel.setSelTile(X, Y, curJ * 19 + curI, undoch)
110                           Call parent.UpdateLevelTile(X, Y, False, True)
120                       End If
130                   Else
140                       If parent.getTile(X, Y) = chJ * 19 + chI Then
150                           Call parent.setTile(X, Y, curJ * 19 + curI, undoch)
160                           Call parent.UpdateLevelTile(X, Y, False, True)
170                       End If
180                   End If
190               End If
200           Next
210       Next
End Sub

Private Sub Swap(ByRef X As Integer, ByRef Y As Integer)
          Dim tmp As Integer
10        tmp = X
20        X = Y
30        Y = tmp

End Sub

Private Sub Class_Initialize()
10        oldposX = -1
20        oldposY = -1
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
20        Set undoch = Nothing
End Sub
