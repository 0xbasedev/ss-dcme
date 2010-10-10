VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Dim TransColor As Long








Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Friend Property Get TransparencyColor() As Long
    TransparencyColor = TransColor
End Property

Friend Property Let TransparencyColor(newcol As Long)
    TransColor = newcol
End Property


Sub DrawTilePreview(tilenr As Integer, highlight As Boolean, tileX As Integer, tileY As Integer)
    'Draws a tile on the preview layer
    
    Dim lbx As Integer, lby As Integer

    If parent.magnifier.UsingPixels Then
    
        lbx = parent.hScr.value \ parent.currenttilew
        lby = parent.vScr.value \ parent.currenttilew
    
        If tilenr = TILE_LRG_ASTEROID Then
            Call DrawRectangle(parent.picPreview.hDC, (tileX - lbx), (tileY - lby), (tileX - lbx + 1), tileY - lby + 1, TilePixelColor(tilenr))
        ElseIf tilenr = TILE_STATION Then
            Call DrawRectangle(parent.picPreview.hDC, (tileX - lbx - 2), (tileY - lby - 2), (tileX - lbx + 3), tileY - lby + 3, TilePixelColor(tilenr))
        ElseIf tilenr = TILE_WORMHOLE Then
            Call DrawRectangle(parent.picPreview.hDC, (tileX - lbx - 2), (tileY - lby - 2), (tileX - lbx + 2), tileY - lby + 2, TilePixelColor(tilenr))
        Else
            Call SetPixel(parent.picPreview.hDC, (tileX - lbx), (tileY - lby), TilePixelColor(tilenr))
        End If
        
    Else
        Call parent.TileRender.DrawTile(tilenr, False, parent.picPreview.hDC, parent.TileToScreenX(tileX), parent.TileToScreenY(tileY), False, True)
    End If
                          
End Sub





Sub DrawTile(tilenr As Integer, highlight As Boolean, hDC As Long, screenX As Integer, screenY As Integer, useTransparentFlyunder As Boolean, drawblack As Boolean, Optional usemapzoom As Boolean = True)
    'Draws a tile somewhere
    
    Dim srcspecialtileshDC As Long
    Dim currenttilew As Integer
    
    currenttilew = IIf(usemapzoom, parent.currenttilew, TILEW)
    
    If tilenr = TILE_WORMHOLE Or tilenr = TILE_LRG_ASTEROID Or tilenr = TILE_STATION Then
        Call DrawObject(tilenr, highlight, hDC, screenX, screenY, 0, 0)

    ElseIf tilenr > 0 Then
        Dim srcZoomTilesethDC As Long
        
        
        If highlight Then    'highlight
            srcZoomTilesethDC = IIf(usemapzoom, parent.picHighlightZoomTileset.hDC, parent.pichighlightTileset.hDC)
        Else
            srcZoomTilesethDC = IIf(usemapzoom, parent.piczoomtileset.hDC, parent.pictileset.hDC)
        End If
        
'        If inselection Then
'            BitBlt hDC, screenX, screenY, currenttilew, currenttilew, srcZoomTilesethDC, ((tilenr - 1) Mod 19) * currenttilew, ((tilenr - 1) \ 19) * currenttilew, vbSrcCopy
'        Else
            If useTransparentFlyunder And tilenr >= TILE_FIRST_FLYUNDER And tilenr <= TILE_LAST_FLYUNDER Then
                
'                If useTransparentFlyunder Then BitBlt hDC, screenX, screenY, currenttilew, currenttilew, parent.pictrans.hDC, 0, 0, vbSrcCopy
                BitBlt hDC, screenX, screenY, currenttilew, currenttilew, parent.pictrans.hDC, 0, 0, vbSrcCopy
                
                TransparentBlt hDC, screenX, screenY, currenttilew, currenttilew, srcZoomTilesethDC, ((tilenr - 1) Mod 19) * currenttilew, ((tilenr - 1) \ 19) * currenttilew, vbBlack
            Else
                BitBlt hDC, screenX, screenY, currenttilew, currenttilew, srcZoomTilesethDC, ((tilenr - 1) Mod 19) * currenttilew, ((tilenr - 1) \ 19) * currenttilew, vbSrcCopy
            End If
            
'        End If
    ElseIf drawblack Then
        'tile 0 ; paint black
        BitBlt hDC, screenX, screenY, currenttilew, currenttilew, parent.picempty.hDC, 0, 0, vbSrcCopy
    End If
End Sub





Sub DrawObject(tilenr As Integer, highlight As Boolean, hDC As Long, screenX As Integer, screenY As Integer, objX As Integer, objY As Integer)
    'objX, objY: Coordinates of the object to draw ; in case of a partial object. Use 0,0 to draw full object
    Dim srcspecialtileshDC As Long
    Dim currenttilew As Integer
    currenttilew = parent.currenttilew
        
    If highlight Then
        srcspecialtileshDC = frmGeneral.pichighlightspecial.hDC
    Else
        srcspecialtileshDC = frmGeneral.picspecial.hDC
    End If
        
    
    SetStretchBltMode hDC, COLORONCOLOR
    
    Dim paintSizeX As Integer, paintSizeY As Integer, objsize As Integer, firstX As Integer
    
    objsize = GetMaxSizeOfObject(tilenr) + 1
    
    paintSizeX = objsize - objX
    paintSizeY = objsize - objY
    
    Select Case tilenr
        Case TILE_LRG_ASTEROID
            firstX = 16
        Case TILE_STATION
            firstX = 128
        Case TILE_WORMHOLE
            firstX = 48
    End Select
    
'    BitBlt hDC, screenX, screenY, paintSizeX * currenttilew, paintSizeY * currenttilew, parent.pictrans.hDC, 0, 0, vbSrcCopy
    
    StretchBlt hDC, screenX, screenY, paintSizeX * currenttilew, paintSizeY * currenttilew, srcspecialtileshDC, firstX + (objX * TILEW), objY * TILEH, paintSizeX * TILEW, paintSizeY * TILEH, vbSrcCopy
    
'        StretchBlt piclevel.hdc, X, Y, currenttilew + currenttilew, currenttilew + currenttilew, srcspecialtileshDC, 16, 0, TILEW + TILEW, TILEW + TILEW, vbSrcCopy
'    ElseIf tilenr = 219 Then
'        If highlight Then srcspecialtileshDC = frmGeneral.pichighlightspecial.hdc Else srcspecialtileshDC = frmGeneral.picspecial.hdc
'        SetStretchBltMode piclevel.hdc, HALFTONE
'        StretchBlt piclevel.hdc, X, Y, 6 * (currenttilew), 6 * (currenttilew), srcspecialtileshDC, 128, 0, 6 * TILEW, 6 * TILEW, vbSrcCopy
'    ElseIf tilenr = 220 Then
'        If highlight Then srcspecialtileshDC = frmGeneral.pichighlightspecial.hdc Else srcspecialtileshDC = frmGeneral.picspecial.hdc
'        SetStretchBltMode piclevel.hdc, HALFTONE
'        StretchBlt piclevel.hdc, X, Y, 5 * (currenttilew), 5 * (currenttilew), srcspecialtileshDC, 48, 0, 5 * TILEW, 5 * TILEW, vbSrcCopy
'
End Sub






Private Sub Class_Terminate()
    Set parent = Nothing
End Sub
