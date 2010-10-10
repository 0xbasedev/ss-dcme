VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain
Dim undoch As Changes

Dim UseAsteroids As Boolean

Dim asteroidTypeCount As Integer
Dim AsteroidToUse(0 To 2) As Integer
    
Dim curSelection As TilesetSelection
Dim airBOptions As airBrushOptions

Private Type airBrushOptions
    BrushSize As Integer
    Density As Double
    TotalItemsToDraw As Long
    tileSize As Coordinate
    isSpecialObject As Boolean
End Type


Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
10        On Error GoTo MouseDown_Error
          
20        If Button <> vbLeftButton And Button <> vbRightButton Then Exit Sub
          
          'get tileset selection
          Dim toolOptions As DrawOptions
          
30        curSelection = parent.tileset.selection(Button)
          
40        Call parent.tline.SetSelection(curSelection)
50        Call parent.tline.SetOptions(toolOptions)
          
          
60        UseAsteroids = CBool(GetSetting("UseAirBrushAsAsteroids", "0"))
          
70        asteroidTypeCount = 0
          
80        If UseAsteroids Then
              'Gather all types of asteroids to use
              
90            If CBool(GetSetting("UseSmallAsteroids1", "0")) Then
100               AsteroidToUse(asteroidTypeCount) = TILE_SML_ASTEROID1
110               asteroidTypeCount = asteroidTypeCount + 1
120           End If
              
130           If CBool(GetSetting("UseSmallAsteroids2", "0")) Then
140               AsteroidToUse(asteroidTypeCount) = TILE_SML_ASTEROID2
150               asteroidTypeCount = asteroidTypeCount + 1
160           End If
              
170           If CBool(GetSetting("UseBigAsteroids", "0")) Then
180               AsteroidToUse(asteroidTypeCount) = TILE_LRG_ASTEROID
190               asteroidTypeCount = asteroidTypeCount + 1
200           End If

              'If no types were checked, forget about asteroids
210           If asteroidTypeCount = 0 Then
220               UseAsteroids = False
230           End If
240       End If
          

          
          Dim decreaseNrOfObjects As Integer 'Factor to reduce the number of
                                         'objects created depending on their size
                                         
250       If Not UseAsteroids And (curSelection.tileSize.X > 1 Or curSelection.tileSize.Y > 1) Then
260           decreaseNrOfObjects = longMaximum(curSelection.tileSize.X, curSelection.tileSize.Y)
270       Else
280           decreaseNrOfObjects = 1
290       End If
          
300       airBOptions.BrushSize = frmGeneral.sldAirbSize.value
310       airBOptions.Density = frmGeneral.sldAirbDensity.value / 100#
          
          'Items to draw = (Area/25) * Density \ decreaseNrOfObjects
320       airBOptions.TotalItemsToDraw = (((airBOptions.BrushSize ^ 2) / 25#) * (airBOptions.Density)) \ decreaseNrOfObjects
330       airBOptions.isSpecialObject = (Not UseAsteroids) And curSelection.isSpecialObject
          
340       If UseAsteroids Then
350           airBOptions.tileSize.X = 0
360           airBOptions.tileSize.Y = 0
370       Else
380           airBOptions.tileSize.X = curSelection.tileSize.X - 1
390           airBOptions.tileSize.Y = curSelection.tileSize.Y - 1
400       End If
              
410       parent.undoredo.ResetRedo
420       Set undoch = New Changes
              
430       Call MouseMove(Button, X, Y)

440       On Error GoTo 0
450       Exit Sub

MouseDown_Error:
460       HandleError Err, "AirBrush.MouseDown"
End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single)
          'On Error GoTo MouseMove_Error

10        If Button <> vbLeftButton And Button <> vbRightButton Then Exit Sub
          
          Dim curtilex As Integer, curtiley As Integer
          
          Dim rndx As Integer, rndy As Integer
          
          Dim maptile As Integer
          
          Dim tilecount As Integer
          Dim i As Integer, j As Integer
          
          Dim tmptile As Integer
          
20        Randomize

setAnotherRandom:

30        curtilex = (parent.Hscr.value + X) \ parent.currenttilew
40        curtiley = (parent.Vscr.value + Y) \ parent.currenttilew
      '    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
      '        Exit Sub
      '    End If




      '        If tmptileSizeX > (frmGeneral.sldAirbSize.Value * 2) Then
      '            tmptileSizeX = (frmGeneral.sldAirbSize.Value * 2) + 1
      '        End If
      '
      '        If tmptileSizeY > (frmGeneral.sldAirbSize.Value * 2) Then
      '            tmptileSizeY = (frmGeneral.sldAirbSize.Value * 2) + 1
      '        End If

          
50        For tilecount = 0 To airBOptions.TotalItemsToDraw
              

60            If UseAsteroids Then
                  Dim rndAsteroid As Integer
                  
70                rndAsteroid = Int((asteroidTypeCount) * Rnd)
                  
80                tmptile = AsteroidToUse(rndAsteroid)
90            Else
100               tmptile = curSelection.tilenr
110           End If
              
120           rndx = Rnd() * Int(airBOptions.BrushSize * 2 - airBOptions.tileSize.X) + curtilex - airBOptions.BrushSize
130           rndy = Rnd() * Int(airBOptions.BrushSize * 2 - airBOptions.tileSize.Y) + curtiley - airBOptions.BrushSize


      '            If airBrushOptions.isSpecialObject Then
      '            ElseIf tmptileSizeX > 1 Or tmptileSizeY > 1 Then
      '                rndx = Rnd() * Int(frmGeneral.sldAirbSize.Value * 2 - (tmptileSizeX - 1)) + curtilex - frmGeneral.sldAirbSize.Value
      '                rndy = Rnd() * Int(frmGeneral.sldAirbSize.Value * 2 - (tmptileSizeY - 1)) + curtiley - frmGeneral.sldAirbSize.Value
      '            Else
      '                rndx = Rnd() * Int(frmGeneral.sldAirbSize.Value * 2) + curtilex - frmGeneral.sldAirbSize.Value
      '                rndy = Rnd() * Int(frmGeneral.sldAirbSize.Value * 2) + curtiley - frmGeneral.sldAirbSize.Value
      '            End If

              'Boundary check
140           If rndx + airBOptions.tileSize.X >= 0 And rndx <= 1023 And rndy + airBOptions.tileSize.Y >= 0 And rndy <= 1023 Then

150               For j = 0 To airBOptions.tileSize.Y
160                   If rndy + j >= 0 And rndy + j <= 1023 Then
                      
170                       For i = 0 To airBOptions.tileSize.X
180                           If rndx + i >= 0 And rndy + j >= 0 Then
      '                            If AreaClearForObject(parent, rndx + i, rndy + j, tmptile) Then
                                  maptile = parent.getTile(rndx + i, rndy + j)
                                  
190                               If parent.pastetype <> p_under Or maptile = 0 Then
200                                   If TileIsSpecial(tmptile) Then
210                                       If AreaClearForObject(parent, rndx + i, rndy + j, tmptile) Then
                                          
220                                           If TileIsSpecial(maptile) Then
230                                               Call SearchAndDestroyObject(parent, rndx + i, rndy + j, undoch, False)
240                                           End If
                                      
250                                           Call setObject(parent, tmptile, rndx, rndy, undoch, False)
260                                       End If
270                                   Else
280                                       If TileIsSpecial(maptile) Then
290                                           Call SearchAndDestroyObject(parent, rndx + i, rndy + j, undoch, False)
300                                       End If

310                                       If curSelection.selectionType = TS_Walltiles Then
320                                           Call parent.walltiles.setIsWallTile(rndx + i, rndy + j)
330                                           tmptile = parent.walltiles.DrawWallTile(rndx + i, rndy + j, True, undoch, False, True)
340                                           Call parent.walltiles.setIsWallTile(rndx + i, rndy + j, True)
350                                       End If
360                                       Call parent.setTile(rndx + i, rndy + j, tmptile + i + 19 * j, undoch)
370                                       Call parent.UpdateLevelTile(rndx + i, rndy + j, False)
                                          
380                                   End If
390                               End If
400                           End If
410                       Next
                          
420                   End If
430               Next
                  
      '                For l = 0 To tmptileSizeY - 1
      '                    For k = 0 To tmptileSizeX - 1
      '                        If rndx + k >= 0 And rndy + l >= 0 And _
      '                           rndx + k < 1024 And rndy + l < 1024 Then
      '                            If AreaClearForObject(parent, rndx + k, rndy + l, tmptile) Then
      '                                'if parent.pastetype <> p_under Or _
      '                                    'parent.pastetype = p_under And parent.getTile(rndx + k, rndy + l) = 0 Then
      '                                If isTileSpecial(parent, rndx + k, rndy + l) Then
      '                                    Call SearchAndDestroyObject(parent, rndx + k, rndy + l, undoch, False)
      '                                End If
      '                                If tmptile = 217 Or tmptile = 219 Or tmptile = 220 Then
      '                                    Call setObject(parent, tmptile, rndx, rndy, undoch, False, False)
      '                                Else
      '                                    If usewalltile Then
      '                                        Call parent.walltiles.setIsWallTile(rndx, rndy)
      '                                        tmptile = parent.walltiles.DrawWallTile(rndx, rndy, True, undoch, False)
      '                                        Call parent.walltiles.setIsWallTile(rndx, rndy, True)
      '                                    End If
      '                                    Call parent.setTile(rndx + k, rndy + l, tmptile + k + 19 * l, undoch)
      '                                    Call parent.UpdateLevelTile(rndx + k, rndy + l, False)
      '                                End If
      '                            End If
      '                        End If
      '                    Next
      '                Next
                  
              'End boundary check
440           End If
              
450       Next

460       If parent.magnifier.UsingPixels Then
470           parent.UpdateLevel
480       Else
490           parent.UpdatePreview
500       End If


510       DoEvents
520       If SharedVar.MouseDown <> 0 Then
530           Sleep 25
540           GoTo setAnotherRandom
550       End If

560       On Error GoTo 0
570       Exit Sub

MouseMove_Error:
580       HandleError Err, "AirBrush.MouseMove"
End Sub

Sub MouseUp()
      'do nothing
10        On Error GoTo MouseUp_Error

20        Call parent.undoredo.AddToUndo(undoch, UNDO_AIRBRUSH)

          'when using tile 0 we need to redraw grid
30        parent.UpdateLevel

40        On Error GoTo 0
50        Exit Sub

MouseUp_Error:
60        HandleError Err, "AirBrush.MouseUp"
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
20        Set undoch = Nothing
End Sub
