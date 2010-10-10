VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TilesetSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OnChange()


Dim parent As frmMain

Public lastButton As Integer

Dim selections(1 To 2) As TilesetSelection

Dim lvzrectangleLastClick As Coordinate



'currently selected tab
Public CurrentTab As TilesetTabs

'Dim curAnimFrame As Long

Friend Property Get selection(button As Integer) As TilesetSelection
    selection = selections(button)
End Property


Sub SetParent(main As frmMain)
    Set parent = main
End Sub

Sub SelectWalltiles(button As Integer, setNr As Integer, setLastButton As Boolean)
    If setNr < 0 Or setNr > 7 Then Exit Sub
    
    selections(button).group = setNr
    
    selections(button).tileSize.X = 1
    selections(button).tileSize.Y = 1
    selections(button).pixelSize.X = TILEW
    selections(button).pixelSize.Y = TILEH
    
    selections(button).selectionType = TS_Walltiles
    
    selections(button).tilenr = parent.walltiles.getWallTile(setNr, 0)
    
    Call parent.walltiles.SetCurwall(setNr)
    
    selections(button).isSpecialObject = False
    


        
   
    If setLastButton Then lastButton = button
    
    
    Call Change
    
    'TOREMOVE---
    
    'Update the tileset preview
    Call frmGeneral.UpdatePreview
    Call parent.UpdatePreview
End Sub


'TOREMOVE---
'Sub RefreshSelections()
''Refreshes display of selected tiles
'    Dim button As Integer
'
'    For button = 1 To 2
'        If selection(button).selectionType = TS_LVZ Then
'            'TODO: Select the correct image in list
'            Call SelectLVZ(button, selections(button).group, selections(button).tilenr, False, 0, 0)
'
'        ElseIf selection(button).selectionType = TS_Walltiles Then
'            'TODO
'            Call SelectWalltiles(button, selections(button).group, False)
'        Else
'            Call SelectTiles(button, selections(button).tilenr, CInt(selections(button).tileSize.X), CInt(selections(button).tileSize.Y), False)
'        End If
'    Next
'End Sub



Sub SelectTiles(button As Integer, tilenr As Integer, w As Integer, h As Integer, setLastButton As Boolean)
'Sets the left selection of the tileset on the given tilenr
    If tilenr = 0 Then Exit Sub
    If tilenr > 256 Then Exit Sub


    
    selections(button).tilenr = tilenr Mod 256
    
    
    'If it's a special object, the correct tilesize must be given
    If tilenr = TILE_LRG_ASTEROID Or tilenr = TILE_STATION Or tilenr = TILE_WORMHOLE Then
        selections(button).isSpecialObject = True
        w = GetMaxSizeOfObject(tilenr) + 1
        h = w
    Else
        selections(button).isSpecialObject = False
    End If
    
    selections(button).tileSize.X = w
    selections(button).tileSize.Y = h
    selections(button).pixelSize.X = w * TILEW
    selections(button).pixelSize.Y = h * TILEW
    
    selections(button).selectionType = TS_Tiles
    



    If setLastButton Then lastButton = button
    
    
    Call Change
    
    'TOREMOVE---
    
    'Update the tileset preview
    Call frmGeneral.UpdatePreview
    Call parent.UpdatePreview
End Sub

Sub SelectLVZ(button As Integer, lvzidx As Integer, imgidx As Integer, setLastButton As Boolean, Optional X As Integer = -1, Optional Y As Integer = -1)
    'If X or Y are not specified, the red selection rectangle in the tileset will simply not be updated
    
    If button = vbLeftButton Or button = vbRightButton Then
    

        
        
        If setLastButton Then lastButton = button
        
        selections(button).group = lvzidx
        selections(button).tilenr = imgidx
        selections(button).selectionType = TS_LVZ
        
        selections(button).pixelSize.X = parent.lvz.getImageWidth(lvzidx, imgidx)
        selections(button).pixelSize.Y = parent.lvz.getImageHeight(lvzidx, imgidx)
        
        If selections(button).pixelSize.X Mod TILEW = 0 Then
            selections(button).tileSize.X = selections(button).pixelSize.X \ TILEW
        Else
            selections(button).tileSize.X = selections(button).pixelSize.X \ TILEW + 1
        End If
        If selections(button).pixelSize.Y Mod TILEH = 0 Then
            selections(button).tileSize.Y = selections(button).pixelSize.Y \ TILEH
        Else
            selections(button).tileSize.Y = selections(button).pixelSize.Y \ TILEH + 1
        End If
        
        If X >= 0 And Y >= 0 Then
            lvzrectangleLastClick.X = X
            lvzrectangleLastClick.Y = Y
        End If

        Call Change
        

    End If
End Sub

Friend Property Get LvzLastClickX() As Integer
    LvzLastClickX = lvzrectangleLastClick.X
End Property

Friend Property Get LvzLastClickY() As Integer
    LvzLastClickY = lvzrectangleLastClick.Y
End Property



'Function SelectLVZAt(button As Integer, X As Integer, Y As Integer, setLastButton As Boolean) As Boolean
'    Dim lvzidx As Integer, imgidx As Integer
'
'    If parent.lvz.getImageDefFromCoordinates(X, Y, lvzidx, imgidx) <> -1 Then
'        Call SelectLVZ(button, lvzidx, imgidx, setLastButton, X, Y)
'
'        SelectLVZAt = True
'    Else
'        SelectLVZAt = False
'    End If
'End Function




Public Sub SwapSelections()
    Dim tmpSel As TilesetSelection
    tmpSel = selections(1)
    selections(1) = selections(2)
    selections(2) = tmpSel
    
    Call Change
    
End Sub



Private Sub Class_Initialize()
    lastButton = vbLeftButton
    CurrentTab = TB_Tiles
    
    
End Sub




Private Sub Change()
    RaiseEvent OnChange
End Sub





Private Sub Class_Terminate()
    Set parent = Nothing
End Sub
