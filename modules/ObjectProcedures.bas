Attribute VB_Name = "ObjectProcedures"
Option Explicit

Sub CompleteObject(ByRef map As frmMain, X As Integer, Y As Integer, undoch As Changes, Optional changeundo As Boolean = True)
      'Completes special tiles
          Dim i As Integer
          Dim j As Integer

10        Select Case map.getTile(X, Y)
          Case 217
20            Call map.setTile(X + 1, Y, -21710, undoch, changeundo)
30            Call map.setTile(X, Y + 1, -21701, undoch, changeundo)
40            Call map.setTile(X + 1, Y + 1, -21711, undoch, changeundo)
50        Case 219
60            For j = Y To Y + 5
70                For i = X To X + 5
80                    Call map.setTile(i, j, -21900 - ((i - X) * 10) - (j - Y), undoch, changeundo)
90                Next
100           Next
110           Call map.setTile(X, Y, 219, undoch, changeundo)
120       Case 220
130           For j = Y To Y + 4
140               For i = X To X + 4
150                   Call map.setTile(i, j, -22000 - ((i - X) * 10) - (j - Y), undoch, changeundo)
160               Next
170           Next
180           Call map.setTile(X, Y, 220, undoch, changeundo)
190       Case Else
              'do nothing
200       End Select
End Sub

Sub CompleteSelObject(ByRef sel As selection, X As Integer, Y As Integer, undoch As Changes, Optional DrawPixel As Boolean = True, Optional appendundo As Boolean = True, Optional setinselection As Boolean = False)
      'Completes special tiles
          Dim i As Integer
          Dim j As Integer

10        Select Case sel.getSelTile(X, Y)
          Case 217
20            Call sel.setSelTile(X + 1, Y, -21710, undoch, DrawPixel, appendundo)
30            Call sel.setSelTile(X, Y + 1, -21701, undoch, DrawPixel, appendundo)
40            Call sel.setSelTile(X + 1, Y + 1, -21711, undoch, DrawPixel, appendundo)

50            If setinselection Then
60                Call sel.setinselection(X + 1, Y, True)
70                Call sel.setinselection(X + 1, Y + 1, True)
80                Call sel.setinselection(X, Y + 1, True)
90            End If

100       Case 219
110           For j = Y To Y + 5
120               For i = X To X + 5
130                   Call sel.setSelTile(i, j, -21900 - ((i - X) * 10) - (j - Y), undoch, DrawPixel, appendundo)
140                   If setinselection Then
150                       Call sel.setinselection(i, j, True)
160                   End If
170               Next
180           Next
190           Call sel.setSelTile(X, Y, 219, undoch, DrawPixel, appendundo)
200       Case 220
210           For j = Y To Y + 4
220               For i = X To X + 4
230                   Call sel.setSelTile(i, j, -22000 - ((i - X) * 10) - (j - Y), undoch, DrawPixel, appendundo)
240                   If setinselection Then
250                       Call sel.setinselection(i, j, True)
260                   End If
270               Next
280           Next
290           Call sel.setSelTile(X, Y, 220, undoch, DrawPixel, appendundo)
300       Case Else
              'do nothing
310       End Select
End Sub

Function SearchObject(ByRef map As frmMain, X As Integer, Y As Integer) As Integer()
          Dim retVal() As Integer
10        ReDim retVal(1)

          Dim j As Integer, i As Integer, tile As Integer
20        tile = map.getTile(X, Y)
30        If tile > 0 Then
40            retVal(0) = X
50            retVal(1) = Y
60            SearchObject = retVal
70            Exit Function
80        Else
90            retVal(0) = X + ((tile Mod 100) \ 10)
100           retVal(1) = Y + tile Mod 10
110           SearchObject = retVal
120       End If

End Function

Sub SearchAndDestroyObject(ByRef map As frmMain, X As Integer, Y As Integer, undoch As Changes, Optional Refresh As Boolean = True, Optional replacewith As Integer = 0)
10        If map.getTile(X, Y) > 0 Then
20            Call DestroyObject(map, X, Y, undoch, Refresh)
30            Exit Sub
40        End If

          Dim obj() As Integer
50        ReDim obj(1)

60        obj = SearchObject(map, X, Y)
70        Call DestroyObject(map, obj(0), obj(1), undoch, Refresh, replacewith)

End Sub

Sub DestroyObject(ByRef map As frmMain, X As Integer, Y As Integer, undoch As Changes, Optional Refresh As Boolean = True, Optional replacewith As Integer = 0)
          Dim i As Integer
          Dim j As Integer

          Dim tilenr As Integer
10        tilenr = map.getTile(X, Y)
          
          Dim maxsize As Integer, maxX As Integer, maxY As Integer
20        maxsize = GetMaxSizeOfObject(tilenr)
30        maxX = X + maxsize
40        maxY = Y + maxsize
          
50        If maxX > 1023 Then maxX = 1023
60        If maxY > 1023 Then maxY = 1023
          
70        For j = Y To maxY
80            For i = X To maxX
90                Call map.setTile(i, j, replacewith, undoch)
100               Call map.UpdateLevelTile(i, j, False)
110           Next
120       Next

130       If Refresh Then
140           map.UpdateLevel
150       End If
End Sub

Function GetMaxSizeOfObject(tilenr) As Integer
10        Select Case tilenr
          Case TILE_LRG_ASTEROID
20            GetMaxSizeOfObject = 1
30        Case TILE_STATION
40            GetMaxSizeOfObject = 5
50        Case TILE_WORMHOLE
60            GetMaxSizeOfObject = 4
70        Case Else
80            GetMaxSizeOfObject = 0
90        End Select
End Function

Function setObject(ByRef map As frmMain, tilenr As Integer, X As Integer, Y As Integer, undoch As Changes, Optional Refresh As Boolean = True) As Boolean
          'Places object on map
          'return true if an object had to be deleted
              
          Dim i As Integer, j As Integer
          Dim objsize As Integer, maxX As Integer, maxY As Integer
          
          
10        objsize = GetMaxSizeOfObject(tilenr)
20        maxX = X + objsize
30        maxY = Y + objsize
40        If maxX > 1023 Then maxX = 1023
50        If maxY > 1023 Then maxY = 1023
          
60        setObject = False
          
70        If Y + objsize > 1023 Or X + objsize > 1023 Or map.getTile(X, Y) = tilenr Then
              'object cant get out of map
80            Exit Function
90        End If

          'first, check if there are any special tiles we need to
          'remove
100       For j = Y To maxY
110           For i = X To maxX
120               If TileIsSpecial(map.getTile(i, j)) Then
                      'we overlap with another, kill the other
130                   Call SearchAndDestroyObject(map, i, j, undoch, False)

                      'warn the line class that an object was found, so it will refresh screen
140                   setObject = True
150               End If
160           Next
170       Next

          'then put the new one

180       For j = Y To maxY
190           For i = X To maxX
200               If i = X And j = Y Then
210                   Call map.setTile(i, j, tilenr, undoch)
220               Else
230                   Call map.setTile(i, j, (-tilenr * 100) - (10 * (i - X)) - (j - Y), undoch)
240               End If
                  
      '            Call map.UpdateLevelTile(i, j, False)
250           Next
260       Next

270       Call map.UpdateLevelObject(X, Y, False, True)
          
280       If Refresh Then map.UpdateLevel
End Function



Function SearchSelObject(ByRef sel As selection, ByVal X As Integer, ByVal Y As Integer) As Integer()
          Dim retVal() As Integer
10        ReDim retVal(1)

          Dim j As Integer, i As Integer, tile As Integer
20        tile = sel.getSelTile(X, Y)
30        If tile > 0 Then
40            retVal(0) = X
50            retVal(1) = Y
60            SearchSelObject = retVal
70        Else
80            retVal(0) = X + ((tile Mod 100) \ 10)
90            retVal(1) = Y + tile Mod 10
100           SearchSelObject = retVal
110       End If
End Function

Sub SearchAndDestroySelObject(sel As selection, X As Integer, Y As Integer, undoch As Changes, Optional replacewith As Integer = 0, Optional appendundo As Boolean = True)
10        If sel.getSelTile(X, Y) > 0 Then
20            Call DestroySelObject(sel, X, Y, undoch, , , appendundo)
30            Exit Sub
40        End If

          Dim obj() As Integer
50        ReDim obj(1)

60        obj = SearchSelObject(sel, X, Y)
70        Call DestroySelObject(sel, obj(0), obj(1), undoch, replacewith, , appendundo)

End Sub

Sub DestroySelObject(sel As selection, X As Integer, Y As Integer, undoch As Changes, Optional replacewith As Integer = 0, Optional RemoveFromSelection As Boolean = False, Optional appendundo As Boolean = True)
          Dim i As Integer, j As Integer
          
          Dim tilenr As Integer
10        tilenr = sel.getSelTile(X, Y)
          
          Dim objsize As Integer, maxX As Integer, maxY As Integer

20        objsize = GetMaxSizeOfObject(tilenr)
30        maxX = X + objsize
40        maxY = Y + objsize
50        If maxX > 1023 Then maxX = 1023
60        If maxY > 1023 Then maxY = 1023
          
70        For j = Y To maxY
80            For i = X To maxX
90                If RemoveFromSelection Then
100                   Call sel.DeleteSelectionTile(i, j, undoch)
110               Else
120                   Call sel.setSelTile(i, j, replacewith, undoch, True, appendundo)
130               End If
140           Next
150       Next
End Sub

Function setSelObject(sel As selection, tilenr As Integer, X As Integer, Y As Integer, undoch As Changes, Optional setinselection As Boolean = True, Optional appendundo As Boolean = True) As Boolean
          Dim i As Integer, j As Integer
          Dim objsize As Integer
10        objsize = GetMaxSizeOfObject(tilenr)
          
20        If Y + objsize > 1023 Or X + objsize > 1023 Or sel.getSelTile(X, Y) = tilenr Then
              'object cant get out of map
30            Exit Function
40        End If

50        setSelObject = False
          
          'first, check if there are any special tiles we need to
          'remove
60        For j = Y To Y + objsize
70            For i = X To X + objsize
80                If TileIsSpecial(sel.getSelTile(i, j)) Then
                      'we overlap with another, kill the other
90                    Call SearchAndDestroySelObject(sel, i, j, undoch, 0, appendundo)
100                   setSelObject = True
110               End If
120           Next
130       Next

          'then put the new one
140       For j = Y To Y + objsize
150           For i = X To X + objsize
160               If i = X And j = Y Then
170                   Call sel.setSelTile(i, j, tilenr, undoch, , appendundo)
180               Else
190                   Call sel.setSelTile(i, j, (-tilenr * 100) - (10 * (i - X)) - (j - Y), undoch, , appendundo)
200               End If
210               If setinselection Then
220                   Call sel.setinselection(i, j, True)
230               End If
240           Next
250       Next
End Function



Function AreaClearForObject(ByRef map As frmMain, ByVal X As Integer, ByVal Y As Integer, ByVal tilenr As Integer) As Boolean
          Dim j As Integer
          Dim i As Integer

10        If (tilenr < 0) Then
20            X = X + ((tilenr Mod 100) \ 10)
30            Y = Y + tilenr Mod 10
40            tilenr = tilenr \ -100
50        End If
                                                  
60        If (tilenr <> 217 And tilenr <> 219 And tilenr <> 220) Or map.pastetype <> p_under Then
70            AreaClearForObject = (map.getTile(X, Y) = 0)
80        Else
90            AreaClearForObject = True
              Dim size As Integer
100           size = GetMaxSizeOfObject(tilenr)
              
110           For j = Y To Y + size
120               If j <= 1023 Then
130                   For i = X To X + size
140                       If i <= 1023 Then
150                           If map.getTile(i, j) <> 0 Then
160                               AreaClearForObject = False
170                               Exit Function
180                           End If
190                       End If
200                   Next i
210               End If
220           Next j
230       End If

End Function

Function AreaClearForMapObject(ByRef map As frmMain, X As Integer, Y As Integer, tilenr As Integer) As Boolean
          Dim j As Integer
          Dim i As Integer

              
10        AreaClearForMapObject = True
20        If (tilenr <> 217 And tilenr <> 219 And tilenr <> 220) Or map.pastetype = p_under Then Exit Function

          Dim size As Integer
30        size = GetMaxSizeOfObject(tilenr)
          
40        For j = Y To Y + size
50            For i = X To X + size
60                If map.sel.getIsInSelection(i, j) = True Then
70                    If map.pastetype = p_normal Or map.sel.getSelTile(i, j) <> 0 Then
80                        AreaClearForMapObject = False
90                        Exit Function
100                   End If
110               End If
120           Next i
130       Next j
End Function

Function SearchObjectNr(ByRef map As frmMain, X As Integer, Y As Integer) As Integer

10        If map.getTile(X, Y) > 0 Then
20            SearchObjectNr = map.getTile(X, Y)
30            Exit Function
40        Else
50            SearchObjectNr = map.getTile(X, Y) \ -100
60        End If
End Function

Function SearchSelObjectNr(sel As selection, X As Integer, Y As Integer) As Integer

10        If sel.getSelTile(X, Y) > 0 Then
20            SearchSelObjectNr = sel.getSelTile(X, Y)
30            Exit Function
40        Else
50            SearchSelObjectNr = sel.getSelTile(X, Y) \ -100
60        End If
End Function
