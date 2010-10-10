VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WallTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    
Dim parent As frmMain

Dim walltiles() As Integer
Dim walltilesetCount As Integer

Dim iswalltile(1023, 1023) As Boolean

Dim tileiswalltile(7, 255) As Boolean

Public curwall As Integer
Public UsingWalltiles As Boolean

Private Sub Class_Initialize()
    'TODO: number of walltile sets is not dynamic at the moment
    
    ReDim walltiles(7, 15) As Integer
    walltilesetCount = 0
    curwall = 0
End Sub

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Property Get GetWalltilesetCount() As Integer
    GetWalltilesetCount = walltilesetCount
End Property

Function AddWalltileset(ByRef tilearray() As Byte, count As Long) As Integer
    'adds new walltile set
    'Returns the id of added set, returns -1 if unable to add
    
    'For now, it only works for 8 sets of 16
    If walltilesetCount >= 8 Or count <> 16 Then
        AddWalltileset = -1
    Else
        Dim i As Integer
        For i = 0 To count - 1
            walltiles(walltilesetCount, i) = CInt(tilearray(i))
        Next
        
        AddWalltileset = walltilesetCount
        walltilesetCount = walltilesetCount + 1
    End If
End Function
Sub ClearWalltileset(id As Integer)
    'removes specified walltile set from memory
    
End Sub
Sub ClearAllWalltilesets()
    'removes all walltile sets from memory
    walltilesetCount = 0
End Sub


Function getWallTile(wallset As Integer, walltile As Integer) As Integer
    getWallTile = walltiles(wallset, walltile)
End Function

Sub setWallTile(wallset As Integer, walltile As Integer, tilenr As Integer)
    walltiles(wallset, walltile) = tilenr
End Sub

Sub setIsWallTile(X As Integer, Y As Integer, Optional reset As Boolean = False)
    iswalltile(X, Y) = Not reset
End Sub

Function getIsWallTile(X As Integer, Y As Integer) As Boolean
    getIsWallTile = iswalltile(X, Y)
End Function
Function tileconvert(tile As Integer) As Integer
'Converts tile position to walltile number
    Select Case tile
    Case 0
        tileconvert = 9
    Case 1
        tileconvert = 13
    Case 2
        tileconvert = 12
    Case 3
        tileconvert = 8
    Case 4
        tileconvert = 11
    Case 5
        tileconvert = 15
    Case 6
        tileconvert = 14
    Case 7
        tileconvert = 10
    Case 8
        tileconvert = 3
    Case 9
        tileconvert = 7
    Case 10
        tileconvert = 6
    Case 11
        tileconvert = 2
    Case 12
        tileconvert = 1
    Case 13
        tileconvert = 5
    Case 14
        tileconvert = 4
    Case 15
        tileconvert = 0
    End Select
End Function

Sub clearWallTile(wallset As Integer)
    Dim i As Integer
    For i = 0 To 15
        walltiles(wallset, i) = 0
    Next
End Sub

Sub SetCurwall(newcurwall As Integer)
    curwall = newcurwall
End Sub



Function IsValidTileforWall(tilenr As Integer) As Boolean
    Dim i As Integer
    For i = 0 To 15
        If walltiles(curwall, i) <> 0 And walltiles(curwall, i) = tilenr Then
            IsValidTileforWall = True
            Exit Function
        End If
    Next
    IsValidTileforWall = False
End Function

Sub SaveWallTiles(path As String)
    On Error GoTo errh

    If Dir$(path) <> "" Then
        Kill path
    End If

    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    Put #f, , walltiles
    Close #f
    
    On Error GoTo 0
    Exit Sub
errh:
    HandleError Err, "Walltiles.SaveWallTiles"
End Sub

Sub LoadWallTiles(path As String)
    On Error GoTo errh
    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    Get #f, , walltiles
    Close #f

    On Error GoTo 0
    Exit Sub
errh:
    HandleError Err, "Walltiles.LoadWallTiles"
End Sub



Private Function CheckWallTile(ByVal X As Integer, ByVal Y As Integer, AdjustNeighbors As Boolean, undoch As Changes, onpreview As Boolean, appendundo As Boolean) As Boolean
        If iswalltile(X, Y) Then
            CheckWallTile = True
            If AdjustNeighbors Then
                If IsValidTileforWall(parent.getTile(X, Y)) And parent.pastetype <> p_under Then
                    Call setMapTile(X, Y, DrawWallTile(X, Y, False, undoch, onpreview, True), undoch, onpreview, appendundo)
                End If
            End If
        ElseIf IsValidTileforWall(parent.getTile(X, Y)) And parent.pastetype <> p_under Then
            CheckWallTile = True
            If AdjustNeighbors Then
                Call setMapTile(X, Y, DrawWallTile(X, Y, False, undoch, onpreview, True), undoch, onpreview, appendundo)
            End If
        Else
            CheckWallTile = False
        End If
End Function

'Function DrawWallTile(X As Integer, Y As Integer, AdjustNeighbors As Boolean = True, Optional undoch As Changes, Optional onpreview As Boolean = True, Optional appendundo As Boolean = True) As Integer
Function DrawWallTile(X As Integer, Y As Integer, AdjustNeighbors As Boolean, undoch As Changes, onpreview As Boolean, appendundo As Boolean) As Integer
    Dim val As Integer
    val = 0
    
    If Not onpreview Then
        val = 0
    
    End If
    
    If X <= 1022 Then
        If CheckWallTile(X + 1, Y, AdjustNeighbors, undoch, onpreview, appendundo) Then val = val + 1
    End If
    
    If Y >= 1 Then
        If CheckWallTile(X, Y - 1, AdjustNeighbors, undoch, onpreview, appendundo) Then val = val + 2
    End If
    
    If X >= 1 Then
        If CheckWallTile(X - 1, Y, AdjustNeighbors, undoch, onpreview, appendundo) Then val = val + 4
    End If
    
    If Y <= 1022 Then
        If CheckWallTile(X, Y + 1, AdjustNeighbors, undoch, onpreview, appendundo) Then val = val + 8
    End If
    
    DrawWallTile = walltiles(curwall, val)


    
    
'    If tmp <> 0 Then
'        DrawWallTile = tmp
'    Else
'        tmp = parent.getTile(X, Y)
'        If tmp > 0 Then
'            DrawWallTile = tmp
'        Else
'            Dim i As Integer
'            For i = 0 To 15
'                tmp = walltiles(curwall, i)
'                If tmp <> 0 Then
'                    DrawWallTile = tmp
'                    Exit Function
'                End If
'            Next
'
'            'There's not a single tile on that walltile set...
'        End If
'    End If
End Function

Function ReplaceWithWalltile(X As Integer, Y As Integer, newwallset As Integer, tilenr As Integer, oldwalltile As Boolean, oldwallset As Integer, replace As Boolean) As Integer
    Dim val As Integer
    Dim tmp As Integer

    val = 0

    If Y >= 1 Then
        If (GetTileIsWalltile(parent.getTile(X, Y - 1), newwallset)) Then val = val + 2
    End If
    If X >= 1 Then
        If (GetTileIsWalltile(parent.getTile(X - 1, Y), newwallset)) Then val = val + 4
    End If
    If Y <= 1022 Then
        If (replace And GetTileIsWalltile(parent.getTile(X, Y + 1), newwallset)) Or parent.getTile(X, Y + 1) = tilenr Or (oldwalltile And tileiswalltile(oldwallset, parent.getTile(X, Y + 1))) Then val = val + 8
    End If
    If X <= 1022 Then
        If (replace And GetTileIsWalltile(parent.getTile(X + 1, Y), newwallset)) Or parent.getTile(X + 1, Y) = tilenr Or (oldwalltile And tileiswalltile(oldwallset, parent.getTile(X + 1, Y))) Then val = val + 1
    End If

    tmp = walltiles(newwallset, val)

    If tmp <> 0 Then
        ReplaceWithWalltile = tmp
    Else
        tmp = parent.getTile(X, Y)
        If tmp > 0 Then
            ReplaceWithWalltile = tmp
        Else
            Dim i As Integer
            For i = 0 To 15
                tmp = walltiles(newwallset, i)
                If tmp <> 0 Then
                    ReplaceWithWalltile = tmp
                    Exit Function
                End If
            Next

            'There's not a single tile on that walltile set...
        End If
    End If

End Function

Function ReplaceSelWithWalltile(X As Integer, Y As Integer, newwallset As Integer, tilenr As Integer, oldwalltile As Boolean, oldwallset As Integer, replace As Boolean) As Integer
    Dim val As Integer
    Dim tmp As Integer

    val = 0

    If Y >= 1 Then
        If (GetTileIsWalltile(parent.sel.getSelTile(X, Y - 1), newwallset)) Then val = val + 2
    End If
    If X >= 1 Then
        If (GetTileIsWalltile(parent.sel.getSelTile(X - 1, Y), newwallset)) Then val = val + 4
    End If
    If Y <= 1022 Then
        If (replace And GetTileIsWalltile(parent.sel.getSelTile(X, Y + 1), newwallset)) Or parent.sel.getSelTile(X, Y + 1) = tilenr Or (oldwalltile And GetTileIsWalltile(parent.sel.getSelTile(X, Y + 1), oldwallset)) Then val = val + 8
    End If
    If X <= 1022 Then
        If (replace And GetTileIsWalltile(parent.sel.getSelTile(X + 1, Y), newwallset)) Or parent.sel.getSelTile(X + 1, Y) = tilenr Or (oldwalltile And GetTileIsWalltile(parent.sel.getSelTile(X + 1, Y), oldwallset)) Then val = val + 1
    End If

    tmp = walltiles(newwallset, val)

    If tmp <> 0 Then
        ReplaceSelWithWalltile = tmp
    Else
        tmp = parent.sel.getSelTile(X, Y)
        If tmp > 0 Then
            ReplaceSelWithWalltile = tmp
        Else
            Dim i As Integer
            For i = 0 To 15
                tmp = walltiles(newwallset, i)
                If tmp <> 0 Then
                    ReplaceSelWithWalltile = tmp
                    Exit Function
                End If
            Next

            'There's not a single tile on that walltile set...
        End If
    End If

End Function


Sub ClearIsWallTile()
    Erase iswalltile
End Sub

Private Sub setMapTile(X As Integer, Y As Integer, tilenr As Integer, undoch As Changes, Optional onpreview As Boolean = False, Optional appendundo As Boolean = True)
    If Not onpreview Then
        Call parent.setTile(X, Y, tilenr, undoch, appendundo)
        Call parent.UpdateLevelTile(X, Y, False)
    Else
    
        Call parent.TileRender.DrawTilePreview(tilenr, False, X, Y)
        
'        Dim lbx As Integer
'        Dim lby As Integer
'        lbx = parent.hScr.value \ parent.currenttilew
'        lby = parent.vScr.value \ parent.currenttilew
'
'        If Not parent.magnifier.UsingPixels Then
'            Dim srcZoomTilesethDC As Long
'            If False Then    'highlight
'                srcZoomTilesethDC = parent.picHighlightZoomTileset.hDC
'            Else
'                srcZoomTilesethDC = parent.piczoomtileset.hDC
'            End If
'            BitBlt parent.picpreview.hDC, (X - lbx) * (parent.currenttilew), (Y - lby) * (parent.currenttilew), (parent.currenttilew), (parent.currenttilew), srcZoomTilesethDC, ((tilenr - 1) Mod 19) * (parent.currenttilew), ((tilenr - 1) \ 19) * (parent.currenttilew), vbSrcCopy
'        Else
'            Call SetPixel(parent.picpreview.hDC, (X - lbx), (Y - lby), TilePixelColor(tilenr))
'        End If
    End If
End Sub

Function DrawSelWallTile(X As Integer, Y As Integer, undoch As Changes, Optional AdjustNeighbors As Boolean = True, Optional onpreview As Boolean = True) As Integer
    Dim val As Integer
    Dim tmp As Integer
    val = 0
    If X <= 1022 Then
        If parent.sel.getIsInSelection(X + 1, Y) And (IsValidTileforWall(parent.sel.getSelTile(X + 1, Y)) And parent.pastetype <> p_under) Or iswalltile(X + 1, Y) Then val = val + 1
        If AdjustNeighbors And parent.pastetype <> p_under And IsValidTileforWall(parent.sel.getSelTile(X + 1, Y)) And Not iswalltile(X + 1, Y) Then
            Call setSelTile(X + 1, Y, DrawSelWallTile(X + 1, Y, undoch, False, onpreview), undoch, onpreview)
        End If
    End If
    If Y >= 1 Then
        If parent.sel.getIsInSelection(X, Y - 1) And (IsValidTileforWall(parent.sel.getSelTile(X, Y - 1)) And parent.pastetype <> p_under) Or iswalltile(X, Y - 1) Then val = val + 2
        If AdjustNeighbors And parent.pastetype <> p_under And IsValidTileforWall(parent.sel.getSelTile(X, Y - 1)) And Not iswalltile(X, Y - 1) Then
            Call setSelTile(X, Y - 1, DrawSelWallTile(X, Y - 1, undoch, False, onpreview), undoch, onpreview)
        End If
    End If
    If X >= 1 Then
        If parent.sel.getIsInSelection(X - 1, Y) And (IsValidTileforWall(parent.sel.getSelTile(X - 1, Y)) And parent.pastetype <> p_under) Or iswalltile(X - 1, Y) Then val = val + 4
        If AdjustNeighbors And parent.pastetype <> p_under And IsValidTileforWall(parent.sel.getSelTile(X - 1, Y)) And Not iswalltile(X - 1, Y) Then
            Call setSelTile(X - 1, Y, DrawSelWallTile(X - 1, Y, undoch, False, onpreview), undoch, onpreview)
        End If
    End If
    If Y <= 1022 Then
        If parent.sel.getIsInSelection(X, Y + 1) And (IsValidTileforWall(parent.sel.getSelTile(X, Y + 1)) And parent.pastetype <> p_under) Or iswalltile(X, Y + 1) Then val = val + 8
        If AdjustNeighbors And parent.pastetype <> p_under And IsValidTileforWall(parent.sel.getSelTile(X, Y + 1)) And Not iswalltile(X, Y + 1) Then
            Call setSelTile(X, Y + 1, DrawSelWallTile(X, Y + 1, undoch, False, onpreview), undoch, onpreview)
        End If
    End If
    tmp = walltiles(curwall, val)

    If tmp <> 0 Then
        DrawSelWallTile = tmp
    Else
        tmp = parent.getTile(X, Y)
        If tmp > 0 Then
            DrawSelWallTile = tmp
        Else
            Dim i As Integer
            For i = 0 To 15
                tmp = walltiles(curwall, i)
                If tmp <> 0 Then
                    DrawSelWallTile = tmp
                    Exit Function
                End If
            Next

            'There's not a single tile on that walltile set...
        End If
    End If
End Function

Private Sub setSelTile(X As Integer, Y As Integer, tilenr As Integer, undoch As Changes, Optional onpreview As Boolean = False)
    If Not onpreview Then
        Call parent.sel.setSelTile(X, Y, tilenr, undoch, True)
    Else
        
        Call parent.TileRender.DrawTilePreview(tilenr, True, X, Y)
    End If
End Sub

Function isValidSet(wallset As Integer) As Boolean
    'Returns true if the walltile set is not empty
    
    Dim i As Integer
    For i = 0 To 15
        If walltiles(wallset, i) <> 0 Then
            isValidSet = True
            Exit Function
        End If
    Next
    isValidSet = False
End Function

Sub SetTileIsWalltile()
    Dim i As Integer
    Dim j As Integer

    Erase tileiswalltile

    For i = 0 To 7
        For j = 0 To 15
            If walltiles(i, j) <> 0 Then
                tileiswalltile(i, walltiles(i, j)) = True
            End If
        Next
    Next

End Sub

Function GetTileIsWalltile(tilenr As Integer, Optional currentset As Integer = -1) As Boolean
    If tilenr < 0 Then tilenr = tilenr \ -100
    If currentset = -1 Then
        'Check if tile is part of any tileset
        Dim i As Integer
        For i = 0 To 7
            If tileiswalltile(i, tilenr) Then
                GetTileIsWalltile = True
                Exit Function
            End If
        Next
    Else
        GetTileIsWalltile = tileiswalltile(currentset, tilenr)
    End If
    
End Function


Sub DrawWallTilesSet(hDC As Long, X As Integer, Y As Integer, setID As Integer)
    'For each wall set, first draw the default image
    BitBlt hDC, X, Y, WT_SETW, WT_SETH, frmGeneral.picdefaultwalltiles.hDC, 0, 0, vbSrcCopy
    
    Dim a As Integer, val As Integer
    
    For a = 0 To 15

        'Get corresponding wall tile
        val = walltiles(setID, tileconvert(a))

        'If a wall tile is set to that spot, draw it
        If val Then
            Call parent.TileRender.DrawTile(val, False, hDC, X + (a Mod WT_NTILESW) * TILEW, Y + (a \ WT_NTILESH) * TILEH, False, True, False)
        End If
    Next
End Sub

Sub DrawWallTiles(hDC As Long, nWidth As Integer)
    'Draws a preview of the walltiles on hDC
    'hDC:     Target hDC
    'nWidth:  Number of sets to draw horizontally

    Dim i As Integer

    For i = 0 To 7
        Call DrawWallTilesSet(hDC, (i Mod nWidth) * WT_SETW, (i \ nWidth) * WT_SETH, i)
    Next

End Sub




Private Sub Class_Terminate()
    Set parent = Nothing
    Erase walltiles
End Sub
