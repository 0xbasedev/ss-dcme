Attribute VB_Name = "ObjectProcedures"
Option Explicit

Sub CompleteObject(ByRef map As frmMain, X As Integer, Y As Integer, undoch As Changes, Optional changeundo As Boolean = True)
'Completes special tiles
    Dim i As Integer
    Dim j As Integer

    Select Case map.getTile(X, Y)
    Case 217
        Call map.setTile(X + 1, Y, -21710, undoch, changeundo)
        Call map.setTile(X, Y + 1, -21701, undoch, changeundo)
        Call map.setTile(X + 1, Y + 1, -21711, undoch, changeundo)
    Case 219
        For j = Y To Y + 5
            For i = X To X + 5
                Call map.setTile(i, j, -21900 - ((i - X) * 10) - (j - Y), undoch, changeundo)
            Next
        Next
        Call map.setTile(X, Y, 219, undoch, changeundo)
    Case 220
        For j = Y To Y + 4
            For i = X To X + 4
                Call map.setTile(i, j, -22000 - ((i - X) * 10) - (j - Y), undoch, changeundo)
            Next
        Next
        Call map.setTile(X, Y, 220, undoch, changeundo)
    Case Else
        'do nothing
    End Select
End Sub

Sub CompleteSelObject(ByRef sel As selection, X As Integer, Y As Integer, undoch As Changes, Optional DrawPixel As Boolean = True, Optional appendundo As Boolean = True, Optional setinselection As Boolean = False)
'Completes special tiles
    Dim i As Integer
    Dim j As Integer

    Select Case sel.getSelTile(X, Y)
    Case 217
        Call sel.setSelTile(X + 1, Y, -21710, undoch, DrawPixel, appendundo)
        Call sel.setSelTile(X, Y + 1, -21701, undoch, DrawPixel, appendundo)
        Call sel.setSelTile(X + 1, Y + 1, -21711, undoch, DrawPixel, appendundo)

        If setinselection Then
            Call sel.setinselection(X + 1, Y, True)
            Call sel.setinselection(X + 1, Y + 1, True)
            Call sel.setinselection(X, Y + 1, True)
        End If

    Case 219
        For j = Y To Y + 5
            For i = X To X + 5
                Call sel.setSelTile(i, j, -21900 - ((i - X) * 10) - (j - Y), undoch, DrawPixel, appendundo)
                If setinselection Then
                    Call sel.setinselection(i, j, True)
                End If
            Next
        Next
        Call sel.setSelTile(X, Y, 219, undoch, DrawPixel, appendundo)
    Case 220
        For j = Y To Y + 4
            For i = X To X + 4
                Call sel.setSelTile(i, j, -22000 - ((i - X) * 10) - (j - Y), undoch, DrawPixel, appendundo)
                If setinselection Then
                    Call sel.setinselection(i, j, True)
                End If
            Next
        Next
        Call sel.setSelTile(X, Y, 220, undoch, DrawPixel, appendundo)
    Case Else
        'do nothing
    End Select
End Sub

Function SearchObject(ByRef map As frmMain, X As Integer, Y As Integer) As Integer()
    Dim retVal() As Integer
    ReDim retVal(1)

    Dim j As Integer, i As Integer, tile As Integer
    tile = map.getTile(X, Y)
    If tile > 0 Then
        retVal(0) = X
        retVal(1) = Y
        SearchObject = retVal
        Exit Function
    Else
        retVal(0) = X + ((tile Mod 100) \ 10)
        retVal(1) = Y + tile Mod 10
        SearchObject = retVal
    End If

End Function

Sub SearchAndDestroyObject(ByRef map As frmMain, X As Integer, Y As Integer, undoch As Changes, Optional Refresh As Boolean = True, Optional replacewith As Integer = 0)
    If map.getTile(X, Y) > 0 Then
        Call DestroyObject(map, X, Y, undoch, Refresh)
        Exit Sub
    End If

    Dim obj() As Integer
    ReDim obj(1)

    obj = SearchObject(map, X, Y)
    Call DestroyObject(map, obj(0), obj(1), undoch, Refresh, replacewith)

End Sub

Sub DestroyObject(ByRef map As frmMain, X As Integer, Y As Integer, undoch As Changes, Optional Refresh As Boolean = True, Optional replacewith As Integer = 0)
    Dim i As Integer
    Dim j As Integer

    Dim tilenr As Integer
    tilenr = map.getTile(X, Y)
    
    Dim maxsize As Integer, maxX As Integer, maxY As Integer
    maxsize = GetMaxSizeOfObject(tilenr)
    maxX = X + maxsize
    maxY = Y + maxsize
    
    If maxX > 1023 Then maxX = 1023
    If maxY > 1023 Then maxY = 1023
    
    For j = Y To maxY
        For i = X To maxX
            Call map.setTile(i, j, replacewith, undoch)
            Call map.UpdateLevelTile(i, j, False)
        Next
    Next

    If Refresh Then
        map.UpdateLevel
    End If
End Sub

Function GetMaxSizeOfObject(tilenr) As Integer
    Select Case tilenr
    Case TILE_LRG_ASTEROID
        GetMaxSizeOfObject = 1
    Case TILE_STATION
        GetMaxSizeOfObject = 5
    Case TILE_WORMHOLE
        GetMaxSizeOfObject = 4
    Case Else
        GetMaxSizeOfObject = 0
    End Select
End Function

Function setObject(ByRef map As frmMain, tilenr As Integer, X As Integer, Y As Integer, undoch As Changes, Optional Refresh As Boolean = True) As Boolean
    'Places object on map
    'return true if an object had to be deleted
        
    Dim i As Integer, j As Integer
    Dim objsize As Integer, maxX As Integer, maxY As Integer
    
    
    objsize = GetMaxSizeOfObject(tilenr)
    maxX = X + objsize
    maxY = Y + objsize
    If maxX > 1023 Then maxX = 1023
    If maxY > 1023 Then maxY = 1023
    
    setObject = False
    
    If Y + objsize > 1023 Or X + objsize > 1023 Or map.getTile(X, Y) = tilenr Then
        'object cant get out of map
        Exit Function
    End If

    'first, check if there are any special tiles we need to
    'remove
    For j = Y To maxY
        For i = X To maxX
            If TileIsSpecial(map.getTile(i, j)) Then
                'we overlap with another, kill the other
                Call SearchAndDestroyObject(map, i, j, undoch, False)

                'warn the line class that an object was found, so it will refresh screen
                setObject = True
            End If
        Next
    Next

    'then put the new one

    For j = Y To maxY
        For i = X To maxX
            If i = X And j = Y Then
                Call map.setTile(i, j, tilenr, undoch)
            Else
                Call map.setTile(i, j, (-tilenr * 100) - (10 * (i - X)) - (j - Y), undoch)
            End If
            
'            Call map.UpdateLevelTile(i, j, False)
        Next
    Next

    Call map.UpdateLevelObject(X, Y, False, True)
    
    If Refresh Then map.UpdateLevel
End Function



Function SearchSelObject(ByRef sel As selection, ByVal X As Integer, ByVal Y As Integer) As Integer()
    Dim retVal() As Integer
    ReDim retVal(1)

    Dim j As Integer, i As Integer, tile As Integer
    tile = sel.getSelTile(X, Y)
    If tile > 0 Then
        retVal(0) = X
        retVal(1) = Y
        SearchSelObject = retVal
    Else
        retVal(0) = X + ((tile Mod 100) \ 10)
        retVal(1) = Y + tile Mod 10
        SearchSelObject = retVal
    End If
End Function

Sub SearchAndDestroySelObject(sel As selection, X As Integer, Y As Integer, undoch As Changes, Optional replacewith As Integer = 0, Optional appendundo As Boolean = True)
    If sel.getSelTile(X, Y) > 0 Then
        Call DestroySelObject(sel, X, Y, undoch, , , appendundo)
        Exit Sub
    End If

    Dim obj() As Integer
    ReDim obj(1)

    obj = SearchSelObject(sel, X, Y)
    Call DestroySelObject(sel, obj(0), obj(1), undoch, replacewith, , appendundo)

End Sub

Sub DestroySelObject(sel As selection, X As Integer, Y As Integer, undoch As Changes, Optional replacewith As Integer = 0, Optional RemoveFromSelection As Boolean = False, Optional appendundo As Boolean = True)
    Dim i As Integer, j As Integer
    
    Dim tilenr As Integer
    tilenr = sel.getSelTile(X, Y)
    
    Dim objsize As Integer, maxX As Integer, maxY As Integer

    objsize = GetMaxSizeOfObject(tilenr)
    maxX = X + objsize
    maxY = Y + objsize
    If maxX > 1023 Then maxX = 1023
    If maxY > 1023 Then maxY = 1023
    
    For j = Y To maxY
        For i = X To maxX
            If RemoveFromSelection Then
                Call sel.DeleteSelectionTile(i, j, undoch)
            Else
                Call sel.setSelTile(i, j, replacewith, undoch, True, appendundo)
            End If
        Next
    Next
End Sub

Function setSelObject(sel As selection, tilenr As Integer, X As Integer, Y As Integer, undoch As Changes, Optional setinselection As Boolean = True, Optional appendundo As Boolean = True) As Boolean
    Dim i As Integer, j As Integer
    Dim objsize As Integer
    objsize = GetMaxSizeOfObject(tilenr)
    
    If Y + objsize > 1023 Or X + objsize > 1023 Or sel.getSelTile(X, Y) = tilenr Then
        'object cant get out of map
        Exit Function
    End If

    setSelObject = False
    
    'first, check if there are any special tiles we need to
    'remove
    For j = Y To Y + objsize
        For i = X To X + objsize
            If TileIsSpecial(sel.getSelTile(i, j)) Then
                'we overlap with another, kill the other
                Call SearchAndDestroySelObject(sel, i, j, undoch, 0, appendundo)
                setSelObject = True
            End If
        Next
    Next

    'then put the new one
    For j = Y To Y + objsize
        For i = X To X + objsize
            If i = X And j = Y Then
                Call sel.setSelTile(i, j, tilenr, undoch, , appendundo)
            Else
                Call sel.setSelTile(i, j, (-tilenr * 100) - (10 * (i - X)) - (j - Y), undoch, , appendundo)
            End If
            If setinselection Then
                Call sel.setinselection(i, j, True)
            End If
        Next
    Next
End Function



Function AreaClearForObject(ByRef map As frmMain, ByVal X As Integer, ByVal Y As Integer, ByVal tilenr As Integer) As Boolean
    Dim j As Integer
    Dim i As Integer

    If (tilenr < 0) Then
        X = X + ((tilenr Mod 100) \ 10)
        Y = Y + tilenr Mod 10
        tilenr = tilenr \ -100
    End If
                                            
    If (tilenr <> 217 And tilenr <> 219 And tilenr <> 220) Or map.pastetype <> p_under Then
        AreaClearForObject = (map.getTile(X, Y) = 0)
    Else
        AreaClearForObject = True
        Dim size As Integer
        size = GetMaxSizeOfObject(tilenr)
        
        For j = Y To Y + size
            If j <= 1023 Then
                For i = X To X + size
                    If i <= 1023 Then
                        If map.getTile(i, j) <> 0 Then
                            AreaClearForObject = False
                            Exit Function
                        End If
                    End If
                Next i
            End If
        Next j
    End If

End Function

Function AreaClearForMapObject(ByRef map As frmMain, X As Integer, Y As Integer, tilenr As Integer) As Boolean
    Dim j As Integer
    Dim i As Integer

        
    AreaClearForMapObject = True
    If (tilenr <> 217 And tilenr <> 219 And tilenr <> 220) Or map.pastetype = p_under Then Exit Function

    Dim size As Integer
    size = GetMaxSizeOfObject(tilenr)
    
    For j = Y To Y + size
        For i = X To X + size
            If map.sel.getIsInSelection(i, j) = True Then
                If map.pastetype = p_normal Or map.sel.getSelTile(i, j) <> 0 Then
                    AreaClearForMapObject = False
                    Exit Function
                End If
            End If
        Next i
    Next j
End Function

Function SearchObjectNr(ByRef map As frmMain, X As Integer, Y As Integer) As Integer

    If map.getTile(X, Y) > 0 Then
        SearchObjectNr = map.getTile(X, Y)
        Exit Function
    Else
        SearchObjectNr = map.getTile(X, Y) \ -100
    End If
End Function

Function SearchSelObjectNr(sel As selection, X As Integer, Y As Integer) As Integer

    If sel.getSelTile(X, Y) > 0 Then
        SearchSelObjectNr = sel.getSelTile(X, Y)
        Exit Function
    Else
        SearchSelObjectNr = sel.getSelTile(X, Y) \ -100
    End If
End Function
