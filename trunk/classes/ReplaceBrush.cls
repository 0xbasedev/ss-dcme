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
    Set parent = Main
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
    If Button = vbLeftButton Then
        sourceButton = vbLeftButton
        replaceButton = vbRightButton
    ElseIf Button = vbRightButton Then
        sourceButton = vbRightButton
        replaceButton = vbLeftButton
    Else
        Exit Sub
    End If
    
    oldposX = -1
    oldposY = -1
    
    srcSelection = parent.tileset.selection(sourceButton)
    newselection = parent.tileset.selection(replaceButton)
    
    parent.undoredo.ResetRedo
    Set undoch = New Changes

    Call MouseMove(Button, X, Y)
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer

    If Button <> vbLeftButton And Button <> vbRightButton Then
        Exit Sub
    End If

    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If curtilex < 0 Then curtilex = 0
    If curtiley < 0 Then curtiley = 0
    If curtilex > 1023 Then curtilex = 1023
    If curtiley > 1023 Then curtiley = 1023

    If oldposX = -1 Or oldposY = -1 Then
        oldposX = curtilex
        oldposY = curtiley
    ElseIf oldposX = curtilex And oldposY = curtiley Then
        Exit Sub
    End If

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
    
    Call DrawReplaceLine(srcSelection.tilenr, newselection.tilenr, oldposX, oldposY, curtilex, curtiley, undoch, CInt(srcSelection.tileSize.X), CInt(srcSelection.tileSize.Y), frmGeneral.toolSize(curtool - 1).value, True, True)

    oldposX = curtilex
    oldposY = curtiley
End Sub

Sub MouseUp()
'do nothing
    If oldposX = -1 Or oldposY = -1 Then Exit Sub

    Call parent.undoredo.AddToUndo(undoch, UNDO_REPLACEBRUSH)
    'for tile 0 being gridded again we need to update the level
    Call parent.UpdateLevel

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
    offset = (width - 1) \ 2

    startx = sX - offset
    startY = sY - offset
    endX = eX - offset
    endY = eY - offset

    If width <> 1 Then
        For j = 0 To width - 1
            For i = 0 To width - 1
                Call DrawBaseLine(srctile, newtile, startx + i, startY + j, endX + i, endY + j, undoch, multTileX, multTileY, False, appendundo)
            Next
        Next
        If Refresh Then
            If parent.magnifier.UsingPixels Then
                parent.UpdateLevel
            Else
                parent.UpdatePreview
            End If
        End If
    Else
        Call DrawBaseLine(srctile, newtile, startx, startY, endX, endY, undoch, multTileX, multTileY, Refresh, appendundo)
    End If

    '    End If
End Sub

Private Sub DrawBaseLine(srctile As Integer, newtile As Integer, startx As Integer, startY As Integer, endX As Integer, endY As Integer, undoch As Changes, multTileX As Integer, multTileY As Integer, Optional Refresh As Boolean = True, Optional appendundo As Boolean = True, Optional tagwalltiles As Boolean = False)
' Bresenham's line algorithm for Microsoft Visual Basic 6.0
' Implementation by Robert Lee <rlee0001@maine.rr.com> July, 2002 Public Domain
    Dim InitialX As Integer
    Dim InitialY As Integer
    Dim FinalX As Integer
    Dim FinalY As Integer

    InitialX = startx
    InitialY = startY
    FinalX = endX
    FinalY = endY

    Dim Steep As Boolean
    Dim DeltaX As Integer, DeltaY As Integer, Delta As Integer
    Dim StepX As Integer, StepY As Integer
    Dim coord As Integer

    Steep = False
    DeltaX = Abs(FinalX - InitialX)
    If (FinalX - InitialX) > 0 Then
        StepX = 1
    Else
        StepX = -1
    End If
    DeltaY = Abs(FinalY - InitialY)
    If (FinalY - InitialY) > 0 Then
        StepY = 1
    Else
        StepY = -1
    End If
    If DeltaY > DeltaX Then
        Steep = True
        Swap InitialX, InitialY
        Swap DeltaX, DeltaY
        Swap StepX, StepY
    End If
    Delta = (DeltaY * 2) - DeltaX
    For coord = 0 To DeltaX - 1
        If Steep Then
            Call DrawTiles(srctile, newtile, InitialY, InitialX, undoch, multTileX, multTileY, appendundo, tagwalltiles)
        Else
            Call DrawTiles(srctile, newtile, InitialX, InitialY, undoch, multTileX, multTileY, appendundo, tagwalltiles)
        End If
        While Delta >= 0
            InitialY = InitialY + StepY
            Delta = Delta - (DeltaX * 2)
        Wend
        InitialX = InitialX + StepX
        Delta = Delta + (DeltaY * 2)
    Next coord

    Call DrawTiles(srctile, newtile, FinalX, FinalY, undoch, multTileX, multTileY, appendundo, tagwalltiles)

    If Refresh Then
        parent.UpdatePreview
    End If
End Sub

Private Sub DrawTiles(srctile As Integer, newtile As Integer, X As Integer, Y As Integer, undoch As Changes, multX As Integer, multY As Integer, Optional appendundo As Boolean = True, Optional tagwalltiles As Boolean = False)
    Dim i As Integer
    Dim j As Integer


    Dim curI As Integer
    Dim curJ As Integer
    Dim chI As Integer
    Dim chJ As Integer
    
    For j = 0 To multY - 1
        curJ = (srctile \ 19) + j
        chJ = (newtile \ 19) + j
        
        For i = 0 To multX - 1

            curI = (srctile Mod 19) + i
            chI = (newtile Mod 19) + i


            If Not (X < 0 Or X > 1023 Or Y < 0 Or Y > 1023) Then
                If parent.sel.getIsInSelection(X, Y) Then
                    If parent.sel.getSelTile(X, Y) = chJ * 19 + chI Then
                        Call parent.sel.setSelTile(X, Y, curJ * 19 + curI, undoch)
                        Call parent.UpdateLevelTile(X, Y, False, True)
                    End If
                Else
                    If parent.getTile(X, Y) = chJ * 19 + chI Then
                        Call parent.setTile(X, Y, curJ * 19 + curI, undoch)
                        Call parent.UpdateLevelTile(X, Y, False, True)
                    End If
                End If
            End If
        Next
    Next
End Sub

Private Sub Swap(ByRef X As Integer, ByRef Y As Integer)
    Dim tmp As Integer
    tmp = X
    X = Y
    Y = tmp

End Sub

Private Sub Class_Initialize()
    oldposX = -1
    oldposY = -1
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
