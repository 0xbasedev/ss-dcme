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
    Select Case tileval
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
        TilePixelColor = vbMagenta
    Case 0    'empty
        TilePixelColor = vbBlack
    Case 1 To 161    'normal
        TilePixelColor = 12632256  'RGB(192, 192, 192)
    Case 162 To 169    'Doors
        TilePixelColor = vbBlue
    Case 170    'flag
        TilePixelColor = vbYellow
    Case 171    'safe
        TilePixelColor = vbGreen
    Case 172    'soccer
        TilePixelColor = vbRed

    Case 173 To 175    'flyover
        TilePixelColor = 160 'RGB(128,128,0) ; dark yellow
    Case 176 To 190    'flyunder
        TilePixelColor = 160  'RGB(160,0,0) ; dark red
    Case 191 To 215
        TilePixelColor = 8388736 'RGB(128,0,128) ; dark magenta
        
    Case 216 To 255
        TilePixelColor = vbMagenta
        
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
    Case Else    'invalid
        TilePixelColor = vbBlack
    End Select
End Function

Function TilesetToolTipText(tilenr As Integer) As String
    Select Case tilenr

    Case 20
        TilesetToolTipText = "Tile " & tilenr & ": Used for map border"
    Case 162 To 165
        TilesetToolTipText = "Tile " & tilenr & ": Door A"
    Case 166 To 169
        TilesetToolTipText = "Tile " & tilenr & ": Door B"
    Case 170
        TilesetToolTipText = "Tile " & tilenr & ": Flag (for turf-style only)"
    Case 171
        TilesetToolTipText = "Tile " & tilenr & ": Safety zone"
    Case 172
        TilesetToolTipText = "Tile " & tilenr & ": Goal tile"
    Case 173 To 175
        TilesetToolTipText = "Tile " & tilenr & ": Fly over tile"
    Case 176 To 190
        TilesetToolTipText = "Tile " & tilenr & ": Fly under tile"
    Case 191
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Ships can go through them, Items bounce off it, Thors go through it. (if you fire an item while in it, the item will float suspended in space)"
    Case 192 To 215
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Solid block (like any other tile)"
    Case 216
        TilesetToolTipText = "Tile " & tilenr & ": Small asteroid #1"
    Case 217
        TilesetToolTipText = "Tile " & tilenr & ": Large asteroid"
    Case 218
        TilesetToolTipText = "Tile " & tilenr & ": Small asteroid #2"
    Case 219
        TilesetToolTipText = "Tile " & tilenr & ": Space station"
    Case 220
        TilesetToolTipText = "Tile " & tilenr & ": Wormhole"
    Case 221 To 240
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Solid block (like any other tile)"
    Case 241
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Visible on radar, Ship can go through it but Items dissapear when they touch it."
    Case 242
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Invisible on radar, Warps your ship on contact, items bounce off it, Thors dissapear on contact"
    Case 243 To 251
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Invisible on radar, Solid block (like any other tile)"
    Case 252
        TilesetToolTipText = "Tile " & tilenr & ": Visible on screen (Animated enemy brick), Invisible on radar, Items go through it, your ship gets warped after a random amount of time (0-2 seconds) while floating on it."
    Case 253
        TilesetToolTipText = "Tile " & tilenr & ": Visible on screen (Animated team brick), Invisible on radar, Items go through it, so does your ship."
    Case 254
        TilesetToolTipText = "Tile " & tilenr & ": Invisible on screen, Invisible on radar. It is impossible to lay bricks while on/near it. Can be used to limit greens spawning area."
    Case 255
        TilesetToolTipText = "Tile " & tilenr & ": Visible On screen (Animated green prize), Invisible on radar, Items go through it, so does your ship. It doesn't show up on radar, you can't pick it up."
    Case 256
        TilesetToolTipText = "Tile 0: Eraser"
    Case Is > 256
        TilesetToolTipText = ""

    Case Else
        TilesetToolTipText = "Tile " & tilenr
    End Select
End Function



Function TileIsSolid(ByVal tilenr As Integer) As Boolean
    If tilenr < 0 Then
        tilenr = tilenr \ -100
    End If

    Select Case tilenr
    Case 0
        TileIsSolid = False
    Case 1 To 169
        TileIsSolid = True
    Case 192 To 219
        TileIsSolid = True
    Case 221 To 240
        TileIsSolid = True
    Case 243 To 251
        TileIsSolid = True
    Case Is < 0
        TileIsSolid = True

    Case Else
        TileIsSolid = False
    End Select
End Function


Function TileIsWarp(tilenr As Integer) As Boolean
    If tilenr < 0 Then
        tilenr = tilenr \ -100
    End If

    Select Case tilenr
    Case 0
        TileIsWarp = False
    Case 220    'wormhole
        TileIsWarp = True
    Case 242    'warp tile
        TileIsWarp = True
    Case 252    'warps after random time
        TileIsWarp = True
    Case Else
        TileIsWarp = False
    End Select
End Function

Function TileIsSpawnable(tilenr As Integer) As Boolean
    If TileIsSolid(tilenr) Then
        TileIsSpawnable = False
    Else
        'some non-solid tiles cannot be spawned on, but i don't feel like testing that now
        TileIsSpawnable = True
    End If

End Function

Function TileIsSpecial(tilenr As Integer) As Boolean
    TileIsSpecial = (tilenr = TILE_LRG_ASTEROID Or _
                     tilenr = TILE_STATION Or _
                     tilenr = TILE_WORMHOLE Or _
                     tilenr < 0)
End Function




