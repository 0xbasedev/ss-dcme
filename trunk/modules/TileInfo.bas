Attribute VB_Name = "TileInfo"
Option Explicit

''''Common tiles


Global Const TILE_FLAG = 170
Global Const TILE_SAFETY = 171
Global Const TILE_GOAL = 172

'Large objects
Global Const TILE_WORMHOLE = 220
Global Const TILE_STATION = 219
Global Const TILE_LRG_ASTEROID = 217

Global Const TILE_SML_ASTEROID1 = 216
Global Const TILE_SML_ASTEROID2 = 218

Global Const TILE_FIRST_FLYUNDER = 176
Global Const TILE_LAST_FLYUNDER = 190







Function TilePixelColor(ByRef tileval As Integer) As Long
10        Select Case tileval
      '    Case -1
      '        'special objects
      '        TilePixelColor = vbMagenta
      '    Case -21799 To -21700
      '        'large asteroid tiles
      '        TilePixelColor = vbMagenta
      '    Case -21999 To -21900
      '        'space station tiles
      '        TilePixelColor = vbMagenta
      '    Case -22099 To -22000
      '        'wormhole tiles
      '        TilePixelColor = vbMagenta
              
          Case Is < 0
              'special objects
20            TilePixelColor = vbMagenta
30        Case 0    'empty
40            TilePixelColor = vbBlack
50        Case 1 To 161    'normal
60            TilePixelColor = 12632256  'RGB(192, 192, 192)
70        Case 162 To 169    'Doors
80            TilePixelColor = vbBlue
90        Case 170    'flag
100           TilePixelColor = vbYellow
110       Case 171    'safe
120           TilePixelColor = vbGreen
130       Case 172    'soccer
140           TilePixelColor = vbRed

150       Case 173 To 175    'flyover
160           TilePixelColor = 160 'RGB(128,128,0) ; dark yellow
170       Case 176 To 190    'flyunder
180           TilePixelColor = 160  'RGB(160,0,0) ; dark red
190       Case 191 To 215
200           TilePixelColor = 8388736 'RGB(128,0,128) ; dark magenta
              
210       Case 216 To 255
220           TilePixelColor = vbMagenta
              
      '    Case 216    'sasteroid1
      '        TilePixelColor = vbMagenta
      '    Case 217    'lasteroid
      '        TilePixelColor = vbMagenta
      '    Case 218    'sasteroid2
      '        TilePixelColor = vbMagenta
      '    Case 219    'starport
      '        TilePixelColor = vbMagenta
      '    Case 220    'wormhole
      '        TilePixelColor = vbMagenta
      '    Case 221 To 255
      '        TilePixelColor = vbMagenta
230       Case Else    'invalid
240           TilePixelColor = vbBlack
250       End Select
End Function

Function TilesetToolTipText(tilenr As Integer) As String
10        Select Case tilenr

          Case 20
20            TilesetToolTipText = "Tile " & tilenr & ": Used for map border"
30        Case 162 To 165
40            TilesetToolTipText = "Tile " & tilenr & ": Door A"
50        Case 166 To 169
60            TilesetToolTipText = "Tile " & tilenr & ": Door B"
70        Case 170
80            TilesetToolTipText = "Tile " & tilenr & ": Flag (for turf-style only)"
90        Case 171
100           TilesetToolTipText = "Tile " & tilenr & ": Safety zone"
110       Case 172
120           TilesetToolTipText = "Tile " & tilenr & ": Goal tile"
130       Case 173 To 175
140           TilesetToolTipText = "Tile " & tilenr & ": Fly over tile"
150       Case 176 To 190
160           TilesetToolTipText = "Tile " & tilenr & ": Fly under tile"
170       Case 191
180           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Ships can go through them, Items bounce off it, Thors go through it. (if you fire an item while in it, the item will float suspended in space)"
190       Case 192 To 215
200           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Solid block (like any other tile)"
210       Case 216
220           TilesetToolTipText = "Tile " & tilenr & ": Small asteroid #1"
230       Case 217
240           TilesetToolTipText = "Tile " & tilenr & ": Large asteroid"
250       Case 218
260           TilesetToolTipText = "Tile " & tilenr & ": Small asteroid #2"
270       Case 219
280           TilesetToolTipText = "Tile " & tilenr & ": Space station"
290       Case 220
300           TilesetToolTipText = "Tile " & tilenr & ": Wormhole"
310       Case 221 To 240
320           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Solid block (like any other tile)"
330       Case 241
340           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Ship can go through it but Items dissapear when they touch it."
350       Case 242
360           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Invisible on radar, Warps your ship on contact, items bounce off it, Thors dissapear on contact"
370       Case 243 To 251
380           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Invisible on radar, Solid block (like any other tile)"
390       Case 252
400           TilesetToolTipText = "Tile " & tilenr & ": Visible on screen (Animated enemy brick), Invisible on radar, Items go through it, your ship gets warped after a random amount of time (0-2 seconds) while floating on it."
410       Case 253
420           TilesetToolTipText = "Tile " & tilenr & ": Visible on screen (Animated team brick), Invisible on radar, Items go through it, so does your ship."
430       Case 254
440           TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Invisible on radar. It is impossible to lay bricks while on/near it. Can be used to limit greens spawning area."
450       Case 255
460           TilesetToolTipText = "Tile " & tilenr & ": Visible On screen (Animated green prize), Invisible on radar, Items go through it, so does your ship. It doesn't show up on radar, you can't pick it up."
470       Case 256
480           TilesetToolTipText = "Tile 0: Eraser"
490       Case Is > 256
500           TilesetToolTipText = ""

510       Case Else
520           TilesetToolTipText = "Tile " & tilenr
530       End Select
End Function



Function TileIsSolid(ByVal tilenr As Integer) As Boolean
10        If tilenr < 0 Then
20            tilenr = tilenr \ -100
30        End If

40        Select Case tilenr
          Case 0
50            TileIsSolid = False
60        Case 1 To 169
70            TileIsSolid = True
80        Case 192 To 219
90            TileIsSolid = True
100       Case 221 To 240
110           TileIsSolid = True
120       Case 243 To 251
130           TileIsSolid = True
140       Case Is < 0
150           TileIsSolid = True

160       Case Else
170           TileIsSolid = False
180       End Select
End Function


Function TileIsWarp(tilenr As Integer) As Boolean
10        If tilenr < 0 Then
20            tilenr = tilenr \ -100
30        End If

40        Select Case tilenr
          Case 0
50            TileIsWarp = False
60        Case 220    'wormhole
70            TileIsWarp = True
80        Case 242    'warp tile
90            TileIsWarp = True
100       Case 252    'warps after random time
110           TileIsWarp = True
120       Case Else
130           TileIsWarp = False
140       End Select
End Function

Function TileIsSpawnable(tilenr As Integer) As Boolean
10        If TileIsSolid(tilenr) Then
20            TileIsSpawnable = False
30        Else
              'some non-solid tiles cannot be spawned on, but i don't feel like testing that now
40            TileIsSpawnable = True
50        End If

End Function

Function TileIsSpecial(tilenr As Integer) As Boolean
10        TileIsSpecial = (tilenr = TILE_LRG_ASTEROID Or _
                           tilenr = TILE_STATION Or _
                           tilenr = TILE_WORMHOLE Or _
                           tilenr < 0)
End Function




