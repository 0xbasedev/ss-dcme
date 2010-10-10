VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'not really the FPS, but the delay between each action loop
'always choose a multiple of 1000
Const FPS = 15
Const FPSTIMEOUT = 5000 'if execution time is missing that many ms, it will just skip them

Const Default_Rotation = 200
Const Default_InitialThrust = 15
Const Default_MaximumThrust = 17
Const Default_InitialSpeed = 2010
Const Default_MaximumSpeed = 3250
Const Default_BounceFactor = 22
Const Default_Recharge = 400
Const Default_MaximumEnergy = 1000
Const Default_WarpRadiusLimit = 20

Const Default_BombFireEnergy = 100
Const Default_BombFireEnergyUpgrade = 100
Const Default_BulletFireEnergy = 20
Const Default_MultiFireEnergy = 20
Const Default_BombInitialLevel = 1
Const Default_BombMaximumLevel = 3
Const Default_BulletInitialLevel = 1
Const Default_BulletMaximumLevel = 3

Const Default_BulletFireDelay = 20
Const Default_MultiFireDelay = 30
    
Dim c_stop As Boolean

Dim ship As ShipStats
Dim shipprops(0 To 7) As ShipProperties

Dim mapprops As MapSettings

Dim drawBombLine As Boolean

Dim isDrawing As Boolean

Dim parent As frmMain

Dim undoch As Changes

Function isRunning() As Boolean
10        isRunning = Not c_stop
End Function

Sub Class_Initialize()
          
10        c_stop = True
          
          'init position to -1 so we can know we need to warp for the first time in StartRun
20        ship.X = -1
30        ship.Y = -1
          
40        ship.vx = 0
50        ship.vy = 0
          
60        ship.ship = CInt(GetSetting("TestMapShip", "1")) - 1
70        ship.freq = 0
End Sub

Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Sub setShipType(i As Integer)
10        ship.ship = i
20        ship.energy = shipprops(ship.ship).MaximumEnergy
30        frmGeneral.optShip(i).value = True
End Sub

Sub StartRun()
         'main loop
          Dim tick As Long
          Dim sleepval As Long
          
10        sleepval = FPS

20        Call ReadSettings
          
30        If ship.X = -1 Or ship.Y = -1 Then
40            Call WarpShip
50        End If
          

          'Add the width of a tile to the pictureboxes so it can be pixel accurate.
          'The updateLevel function will need to work on one more tile for smooth display

          
60        With parent.picPreview
      '        .width = parent.picPreview.width + parent.currenttilew
      '        .height = parent.picPreview.height + parent.currenttilew
      '
      '        parent.piclevel.width = .width
      '        parent.piclevel.height = .height
      '
      '        parent.piclvzunder.width = .width
      '        parent.piclvzunder.height = .height
      '
      '        parent.piclvzover.width = .width
      '        parent.piclvzover.height = .height
70        End With
          
80        parent.shpcursor.visible = False
          
          'used for debugging, real-time vs execution time
          Dim starttick As Long
          'used to calculate frame per second
          Dim displayedframes As Integer
          Dim cyclescount As Long
          
          'set zoom to normal
90        If parent.magnifier.zoom <> 1 Then
100           Call parent.magnifier.SetZoom(1#, False)
110       End If
          
120       GetKeys
          
130       c_stop = False
          
140       ship.energy = shipprops(ship.ship).MaximumEnergy
          
150       parent.UpdateLevel
160       frmGeneral.UpdateToolBarButtons
          
          Dim lbx As Integer    'lowestval for x
          Dim hbx As Integer    'highestval for x
          Dim lby As Integer    'lowestval for y
          Dim hby As Integer    'highestval for y
          
170       Do While Not c_stop
180           tick = GetTickCount
190           While GetTickCount - tick < sleepval
                  'Debug.Print "waiting"
200               Sleep 1
210               DoEvents
220           Wend
              
230           frmGeneral.Label6.visible = True
240           frmGeneral.Label7.visible = True
              
250           frmGeneral.lblposition.Caption = Format(ship.X, "0.0") & " - " & Format(ship.Y, "0.0")
260           frmGeneral.Label7.Caption = " Speed: " & ship.vx & " ; " & ship.vy & " (" & Format(Sqr(ship.vx * ship.vx + ship.vy * ship.vy), "0.00") & ") Aim:" & RoundAimAngle * 180# / PI & Chr(176) & " Energy:" & Int(ship.energy)
270           DoEvents
              
              
280           GetKeys
              
              If c_stop Then Exit Sub
              
290           CollisionCheck
              
300           UpdateShipPos
              
310           Call setMapPosFromShipPos
              
320           CheckShipWarp

330           If ship.energy < shipprops(ship.ship).MaximumEnergy Then
340               If ship.energy + shipprops(ship.ship).Recharge >= shipprops(ship.ship).MaximumEnergy Then
350                   ship.energy = shipprops(ship.ship).MaximumEnergy
360               Else
370                   ship.energy = ship.energy + shipprops(ship.ship).Recharge
380               End If
390           End If
              
400           If sleepval > 0 Then
410               Call parent.UpdateLevel(True, False)

420               Call parent.UpdatePreview(False, True)


430               lbx = parent.hScr.value \ parent.currenttilew
440               hbx = (parent.hScr.value + parent.picPreview.width) \ parent.currenttilew
450               lby = parent.vScr.value \ parent.currenttilew
460               hby = (parent.vScr.value + parent.picPreview.height) \ parent.currenttilew
                  
              '    frmGeneral.Label6.Caption = ship.Y
                  Dim offsetX As Integer
                  Dim offsetY As Integer
                  
470               If Int(ship.X) = lbx + (hbx - lbx) \ 2 And _
                     Int(ship.Y) = lby + (hby - lby) \ 2 Then
480                   offsetX = -(Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
490                   offsetY = -(Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)
500               ElseIf Int(ship.X) = lbx + (hbx - lbx) \ 2 Then
510                   offsetX = -(Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
520                   offsetY = 0
530               ElseIf Int(ship.Y) = lby + (hby - lby) \ 2 Then
540                   offsetX = 0
550                   offsetY = -(Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)
560               Else
570                   offsetX = 0
580                   offsetY = 0
590               End If
                  
                  Call parent.UpdatePreview(False, False)
      '            Call BitBlt(parent.picPreview.hDC, offsetX, offsetY, parent.piclevel.width + parent.currenttilew, parent.piclevel.height + parent.currenttilew, parent.piclevel.hDC, 0, 0, vbSrcCopy)
                  
      '            If parent.ShowLVZ Then
      '                Call parent.UpdateLVZPreview(lbx, lby, hbx, hby, True)
      '            End If
      '
600               Call drawShip
                  
610               If drawBombLine Then
                      'Dim t As Long
                      't = GetTickCount
620                   Call drawBombPath
                      
630               End If
                  
                  'to calculate FPS properly
640               displayedframes = displayedframes + 1
                  
650               parent.picPreview.Refresh

660           ElseIf sleepval < -FPSTIMEOUT Then
                  'something is wrong, 5 seconds are missing... just forget them (happens often in the IDE,
                  'when pausing execution.. could also happen if the application is temporarly stopped for some reason
670               sleepval = 0
680           Else
                  'we got to skip some frames
690           End If
              
              'calculate FPS every second
              'FPS  -  real time / desired time
700           If cyclescount * FPS >= 1000 Then
710               frmGeneral.Label6.Caption = "FPS: " & displayedframes & " - " & GetTickCount - starttick & "ms / " & cyclescount * FPS & "ms"
720               starttick = GetTickCount
730               displayedframes = 0
740               cyclescount = 0
750           End If
              
760           cyclescount = cyclescount + 1
              
770           sleepval = FPS - (GetTickCount - tick) + sleepval
780       Loop
          
          'call getkeys again in case user is still holding Q or W, it will end the drawing and append to undo if needed
790       GetKeys
          
800       Call parent.UpdatePreview
         
End Sub

Sub StopRun()
10        c_stop = True
          
20        If Not bDEBUG Then
30            frmGeneral.Label6.visible = False
40            frmGeneral.Label7.visible = False
50        End If

60        If isDrawing Then
              'we were drawing tiles before, append to undo
70            Call parent.undoredo.AddToUndo(undoch, UNDO_SHIPDRAW)
80            isDrawing = False
90        End If
          
          'update preview picture
100       parent.picPreview.width = (parent.ScaleWidth - parent.vScr.width)
110       parent.picPreview.height = (parent.ScaleHeight - parent.hScr.height)

          'update level picture according to preview picture
120       With parent.picPreview
      '        parent.piclevel.width = .width
      '        parent.piclevel.height = .height
      '        parent.piclvzunder.width = .width
      '        parent.piclvzunder.height = .height
      '        parent.piclvzover.width = .width
      '        parent.piclvzover.height = .height
130       End With
              

          
140       parent.UpdateLevel
150       frmGeneral.UpdateToolBarButtons
              
End Sub

Sub GetKeys()

10        If Not isRunning Then
20            If isDrawing Then
                  'we were drawing tiles before, append to undo
30                Call parent.undoredo.AddToUndo(undoch, UNDO_SHIPDRAW)
40            End If
50            GetAsyncKeyState vbKeyQ
60            GetAsyncKeyState vbKeyW
70            isDrawing = False
80        Else
          
90            If GetAsyncKeyState(vbKeyShift) <> 0 And (GetAsyncKeyState(vbKeyUp) <> 0 Or GetAsyncKeyState(vbKeyDown) <> 0) Then
100               ship.turbo = True
110           Else
120               ship.turbo = False
130           End If
              
140           If GetAsyncKeyState(vbKeyUp) <> 0 Then
150               KeyPress (vbKeyUp)
160           ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
170               KeyPress (vbKeyDown)
180           Else
190               Call ChangeSpeed(False, False, False) 'cancel any remaining afterburner
200           End If
210           If GetAsyncKeyState(vbKeyLeft) <> 0 Then
220               KeyPress (vbKeyLeft)
230           End If
240           If GetAsyncKeyState(vbKeyRight) <> 0 Then
250               KeyPress (vbKeyRight)
260           End If
270           If GetAsyncKeyState(vbKeyEscape) <> 0 Then
280               KeyPress (vbKeyEscape)
290           End If
300           If GetAsyncKeyState(vbKeyInsert) <> 0 Then
310               KeyPress (vbKeyInsert)
320           End If
330           If GetAsyncKeyState(vbKeyB) <> 0 Then
340               KeyPress (vbKeyB)
350           End If
360           If GetAsyncKeyState(vbKeyTab) <> 0 Then
370               KeyPress (vbKeyTab)
380           End If
390           If GetAsyncKeyState(vbKeyControl) <> 0 Then
400               KeyPress (vbKeyControl)
410           End If
              
420           If GetAsyncKeyState(vbKeyQ) <> 0 Then
430               If Not isDrawing Then
440                   Set undoch = New Changes
450                   parent.undoredo.ResetRedo
460                   isDrawing = True
470               End If
480               KeyPress (vbKeyQ)
490           ElseIf GetAsyncKeyState(vbKeyW) <> 0 Then
500               If Not isDrawing Then
510                   Set undoch = New Changes
520                   parent.undoredo.ResetRedo
530                   isDrawing = True
540               End If
550               KeyPress (vbKeyW)
560           Else
                  'neither Q or W are pressed
570               If isDrawing Then
                      'we were drawing tiles before, append to undo
580                   Call parent.undoredo.AddToUndo(undoch, UNDO_SHIPDRAW)
590               End If

600               isDrawing = False
610           End If
              
              Dim i As Integer
620           For i = 0 To 7
630               If GetAsyncKeyState(vbKey1 + i) <> 0 Then
640                   KeyPress (vbKey1 + i)
650                   Exit For
660               End If
670           Next
              
680       End If
          
End Sub


Sub KeyPress(KeyAscii As Integer)
10        If Not isRunning Then Exit Sub
          
20        drawBombLine = False
          
30        Select Case KeyAscii
              Case vbKeyUp

40                Call ChangeSpeed(True, False, False)
              
50            Case vbKeyDown
                  
60                Call ChangeSpeed(False, True, False)
              
70            Case vbKeyRight
80                ship.aimangle = ship.aimangle + shipprops(ship.ship).Rotation
90                If ship.aimangle > 2 * PI Then ship.aimangle = ship.aimangle - 2 * PI
                  
100           Case vbKeyLeft
110               ship.aimangle = ship.aimangle - shipprops(ship.ship).Rotation
120               If ship.aimangle < 0 Then ship.aimangle = ship.aimangle + 2 * PI
              
130           Case vbKeyInsert
140               If ship.energy = shipprops(ship.ship).MaximumEnergy Then
150                   Call WarpShip
160               End If
170           Case vbKeyEscape
180               StopRun
              
190           Case vbKeyB
200               drawBombLine = True
                  
210           Case vbKeyQ
220               If parent.tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
230                   Call parent.setTile(Int(ship.X), Int(ship.Y), parent.tileset.selection(vbLeftButton).tilenr, undoch, True)
240                   Call parent.UpdateLevelTile(Int(ship.X), Int(ship.Y), False, True)
250               End If
260           Case vbKeyW
270               If parent.tileset.selection(vbRightButton).selectionType = TS_Tiles Then
280                   Call parent.setTile(Int(ship.X), Int(ship.Y), parent.tileset.selection(vbRightButton).tilenr, undoch, True)
290                   Call parent.UpdateLevelTile(Int(ship.X), Int(ship.Y), False, True)
300               End If
310           Case vbKey1 To vbKey8
320               If ship.energy = shipprops(ship.ship).MaximumEnergy Then
330                   Call setShipType(KeyAscii - vbKey1)
340               End If
              
350           Case vbKeyTab
360               FireBomb
370           Case vbKeyControl
380               FireBullet
                  
390       End Select
End Sub

Private Sub CollisionCheck()
    On Error GoTo errh:
          
10        If Not isRunning Then Exit Sub
          
          'calculate tile left of ship
          Dim testTileX As Integer
          Dim testTileY As Integer
          Dim i As Integer
          Dim j As Integer
          
          'LEFT CHECK
          '-----------
20        If Int(ship.X + ship.vx * FPS / (10000# * TILEW)) < Int(ship.X) Then
              'we're about to hit the wall
30                For i = 0 To (shipprops(ship.ship).Ysize \ TILEW) - 1
40                    testTileY = Round(ship.Y + (((shipprops(ship.ship).Ysize Mod TILEW) / 2) / TILEW)) + i 'Int(ship.X + (((shipprops(ship.ship).Xsize Mod TileW) / 2) / TileW)) + i
50                    testTileX = Int(ship.X) - 1

                      'check indices first
60                    If testTileX < 0 Then
70                        Call BounceX
80                        GoTo continue1
90                    ElseIf testTileY < 0 Then
100                       Call BounceY
110                   Else
                          'we have a valid x
120                       If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
130                           Call BounceX
140                           GoTo continue1
150                       End If
160                   End If
170               Next
180       End If
          
continue1:
          'TOP CHECK
          '------------
190       If Int(ship.Y + ship.vy * FPS / (10000# * TILEW)) < Int(ship.Y) Then
              'we're about to hit the wall
200               For i = 0 To (shipprops(ship.ship).Xsize \ TILEW) - 1
210                   testTileX = Round(ship.X + (((shipprops(ship.ship).Xsize Mod TILEW) / 2) / TILEW)) + i 'Int(ship.X + (((shipprops(ship.ship).Xsize Mod TileW) / 2) / TileW)) + i
220                   testTileY = Int(ship.Y) - 1

                      'check indices first
230                   If testTileY < 0 Then
240                       Call BounceY
250                       GoTo continue2
260                   Else
                          'we have a valid y
270                       If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
280                           ship.Y = testTileY + 1.05
290                           Call BounceY
300                           GoTo continue2
310                       End If
320                   End If
330               Next
340       End If

continue2:
          'RIGHT CHECK
          '-----------
350       If Int(ship.X + ship.vx * FPS / (10000# * TILEW)) > Int(ship.X) Then
              'we're about to hit the wall
360               For i = 0 To (shipprops(ship.ship).Ysize \ TILEW) - 1
370                   testTileX = Round(ship.X + (shipprops(ship.ship).Xsize / TILEW))
380                   testTileY = Int(ship.Y + (((shipprops(ship.ship).Ysize Mod TILEW) / 2) / TILEW)) + i
                  
                      'check indices first
390                   If testTileX > 1023 Then
400                       Call BounceX
410                       GoTo continue3
420                   Else
                          'we have a valid x
430                       If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
440                           Call BounceX
450                           GoTo continue3
460                       End If
470                   End If
480               Next
490       End If

continue3:
          'BOTTOM CHECK
          '------------
500       If Int(ship.Y + ship.vy * FPS / (10000# * TILEW)) > Int(ship.Y) Then
              'we're about to hit the wall
510               For i = 0 To (shipprops(ship.ship).Xsize \ TILEW) - 1
520                   testTileX = Round(ship.X + (((shipprops(ship.ship).Xsize Mod TILEW) / 2) / TILEW)) + i
530                   testTileY = Int(ship.Y + (shipprops(ship.ship).Ysize \ TILEW) - (((shipprops(ship.ship).Ysize Mod TILEW) / 2) / TILEW)) + 1
                  
                      'check indices first
540                   If testTileY > 1023 Then
550                       Call BounceY
560                       GoTo continue4
570                   Else
                          'we have a valid y
580                       If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
590                           Call BounceY
600                           GoTo continue4
610                       End If
620                   End If
630               Next
640       End If
          
continue4:

650   Exit Sub
errh:

End Sub

Private Sub BounceX()
10        If Not isRunning Then Exit Sub
          
20        ship.vx = -ship.vx * mapprops.BounceFactor
30        ship.vy = ship.vy * mapprops.BounceFactor
End Sub

Private Sub BounceY()
10        If Not isRunning Then Exit Sub
          
20        ship.vy = -ship.vy * mapprops.BounceFactor
30        ship.vx = ship.vx * mapprops.BounceFactor
End Sub

Private Sub UpdateShipPos()
10        If Not isRunning Then Exit Sub
          
20        ship.X = ship.X + ship.vx * FPS / (10000# * TILEW)
30        ship.Y = ship.Y + ship.vy * FPS / (10000# * TILEW)
End Sub

Private Sub setMapPosFromShipPos()
10        On Error GoTo errh
          
20        If Not isRunning Then Exit Sub
          
          Dim lbx As Integer    'lowestval for x
          Dim hbx As Integer    'highestval for x
          Dim lby As Integer    'lowestval for y
          Dim hby As Integer    'highestval for y
30        lbx = parent.hScr.value
40        hbx = (parent.hScr.value + parent.picPreview.width)
50        lby = parent.vScr.value
60        hby = (parent.vScr.value + parent.picPreview.height)
              
          Dim xpos As Single
          Dim ypos As Single
70        xpos = ship.X - ((hbx - lbx) / parent.currenttilew) / 2
80        ypos = ship.Y - ((hby - lby) / parent.currenttilew) / 2
          
'90        If xpos < 0 Then xpos = 0
'100       If ypos < 0 Then ypos = 0
'110       If xpos + (hbx - lbx) / parent.currenttilew > 1023 Then xpos = 1023 - (hbx - lbx - 1)
'120       If ypos + (hby - lby) / parent.currenttilew > 1023 Then ypos = 1023 - (hby - lby - 1)
          

130       Call parent.SetScrollbarValues(xpos * parent.currenttilew, ypos * parent.currenttilew, False)

          
140       Exit Sub
errh:
          
End Sub

Private Sub drawShip()
10        If Not isRunning Then Exit Sub
          
          Dim lbx As Integer    'lowestval for x
          Dim hbx As Integer    'highestval for x
          Dim lby As Integer    'lowestval for y
          Dim hby As Integer    'highestval for y
20        lbx = parent.hScr.value
30        hbx = (parent.hScr.value + parent.picPreview.width)
40        lby = parent.vScr.value
50        hby = (parent.vScr.value + parent.picPreview.height)

          Dim xSrc As Integer
          Dim ySrc As Integer
          Dim shpAngD As Integer
60        shpAngD = 90 + ((RoundAimAngle / PI) * 180)
70        shpAngD = shpAngD Mod 360
          
80        xSrc = Int(((shpAngD Mod 90) / 9) * (shipprops(ship.ship).Xsize) / shipprops(ship.ship).Xsize) * shipprops(ship.ship).Xsize
90        ySrc = Int((shpAngD / 360) * 4) * shipprops(ship.ship).Ysize ' + _
                  'ship.ship * (shipprops(ship.ship).Ysize * 4) '0 ' Int((PI / 4) / ship.angle * shipprops(ship.ship).Ysize)
          
          Call TransparentBlt(parent.picPreview.hDC, Int(ship.X * parent.currenttilew - lbx), Int(ship.Y * parent.currenttilew - lby), shipprops(ship.ship).Xsize, shipprops(ship.ship).Ysize, frmGeneral.picShip(ship.ship).hDC, xSrc, ySrc, vbBlack)
          
'          Exit Sub
'100       If Int(ship.X) = lbx + (hbx - lbx) \ 2 And _
'             Int(ship.Y) = lby + (hby - lby) \ 2 Then
'110           Call TransparentBlt(parent.picPreview.hDC, (Int(ship.X) - lbx) * (parent.currenttilew), (Int(ship.Y) - lby) * (parent.currenttilew), shipprops(ship.ship).Xsize, shipprops(ship.ship).Ysize, frmGeneral.picShip(ship.ship).hDC, xSrc, ySrc, vbBlack)
'120       ElseIf Int(ship.X) = lbx + (hbx - lbx) \ 2 Then
'130           Call TransparentBlt(parent.picPreview.hDC, (Int(ship.X) - lbx) * (parent.currenttilew), ((ship.Y) - lby) * (parent.currenttilew), shipprops(ship.ship).Xsize, shipprops(ship.ship).Ysize, frmGeneral.picShip(ship.ship).hDC, xSrc, ySrc, vbBlack)
'140       ElseIf Int(ship.Y) = lby + (hby - lby) \ 2 Then
'150           Call TransparentBlt(parent.picPreview.hDC, ((ship.X) - lbx) * (parent.currenttilew), (Int(ship.Y) - lby) * (parent.currenttilew), shipprops(ship.ship).Xsize, shipprops(ship.ship).Ysize, frmGeneral.picShip(ship.ship).hDC, xSrc, ySrc, vbBlack)
'160       Else
'170           Call TransparentBlt(parent.picPreview.hDC, ((ship.X) - lbx) * (parent.currenttilew), ((ship.Y) - lby) * (parent.currenttilew), shipprops(ship.ship).Xsize, shipprops(ship.ship).Ysize, frmGeneral.picShip(ship.ship).hDC, xSrc, ySrc, vbBlack)
'180       End If
End Sub


Private Sub CheckShipWarp()
          'checks tiles under ship, and warps ship if a warp tile is found
          Dim i As Integer
          Dim j As Integer
          
10        If frmGeneral.chkTileCollision.value = checked Then
20            For i = Int(ship.X - shipprops(ship.ship).radius) To RoundAway(ship.X + shipprops(ship.ship).radius)
30                For j = Int(ship.Y - shipprops(ship.ship).radius) To RoundAway(ship.Y + shipprops(ship.ship).radius)
40                    If i >= 0 And i <= 1023 And j >= 0 And j <= 1023 Then
50                        If TileIsWarp(parent.getTile(i, j)) Then
60                            Call WarpShip
70                            Exit Sub
80                        End If
90                    End If
100               Next
110           Next
120       End If
End Sub


Sub WarpShip(Optional warpX As Integer = -1, Optional warpY As Integer = -1)
          Dim toX As Integer
          Dim toY As Integer
          Dim tries As Long
10        tries = 1
          
20        Randomize
30        Do
40            If warpX = -1 Then
50                toX = CInt(Rnd() * (mapprops.SpawnRadius(ship.freq Mod 4) * 2 + 1)) + mapprops.SpawnX(ship.freq Mod 4)
60            Else
70                toX = warpX
80            End If
90            If warpY < 0 Or warpY > 1023 Then
100               toY = CInt(Rnd() * (mapprops.SpawnRadius(ship.freq Mod 4) * 2 + 1)) + mapprops.SpawnY(ship.freq Mod 4)
110           Else
120               toY = warpY
130           End If
              
140           If toX < 0 Then toX = 0
150           If toX > 1023 Then toX = 1023
160           If toY < 0 Then toY = 0
170           If toY > 1023 Then toY = 1023
              
180           tries = tries + 1

          'try coordinates until tile is free, abort after 10000 tries
190       Loop While (TileIsSolid(parent.getTile(toX, toY)) And tries < 10000)
          '<<<<<<<<<<<<<<<<<<<<<<<<
          'TODO: SHIP MUST HAVE ENOUGH SPACE AROUND IT, NOT ONLY AT THE TILE
          
200       ship.energy = 0
          
210       ship.X = toX
220       ship.Y = toY
230       ship.vx = 0
240       ship.vy = 0
          
          'randomize orientation
250       ship.aimangle = 2 * PI * Rnd()

End Sub


Private Sub drawBombPath()
          Dim lbx As Integer    'lowestval for x
         ' Dim hbx As Integer    'highestval for x
          Dim lby As Integer    'lowestval for y
         ' Dim hby As Integer    'highestval for y
10        lbx = parent.hScr.value / parent.currenttilew
         ' hbx = (parent.Hscr.value + parent.picpreview.Width) \ parent.currenttilew
20        lby = parent.vScr.value / parent.currenttilew
         ' hby = (parent.Vscr.value + parent.picpreview.Height) \ parent.currenttilew


          Dim val() As Double
          
          Dim vx As Double
          Dim vy As Double
30        vy = Sin(RoundAimAngle)
40        vx = Cos(RoundAimAngle)
          
50        val = getNextIntersecion(Int(ship.X) + (shipprops(ship.ship).Xsize / TILEW) / 2, Int(ship.Y) + (shipprops(ship.ship).Ysize / TILEW) / 2, vx, vy) '2 * (ship.vx / shipprops(ship.ship).MaxSpeed), 2 * (ship.vy / shipprops(ship.ship).MaxSpeed))
          
          Dim SrcX As Integer
          Dim SrcY As Integer
          Dim dstx As Integer
          Dim dsty As Integer
          
60        SrcX = Int(ship.X - lbx) * TILEW + (shipprops(ship.ship).Xsize / 2) '+ (Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
70        SrcY = Int(ship.Y - lby) * TILEW + (shipprops(ship.ship).Ysize / 2) '+ (Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)

          
80        dstx = ((val(0) - lbx) * TILEW) - (Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
90        dsty = ((val(1) - lby) * TILEW) - (Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)
          
       '   If srcX < 0 Then
       '       srcX = 0
       '       srcY =
       '   End If
         ' If srcX > parent.picpreview.ScaleWidth Then
         '     srcX = parent.picpreview.ScaleWidth
         ' End If
         '
         ' If srcY < 0 Then
         '     srcY = 0
         '     srcX = -srcY * (dstX - srcX) / (dstY - srcY) + srcX
         ' End If
         ' If srcY > parent.picpreview.ScaleHeight Then
         '     srcY = parent.picpreview.ScaleHeight
         ' End If
          
          'If dstX < 0 Then
          '    dstX = 0
          'End If
         ' If dstX > parent.picpreview.ScaleWidth Then
         '     dstX = parent.picpreview.ScaleWidth
         ' End If
          
          'If dstY < 0 Then
          '    dstY = 0
          'End If
         ' If dstY > parent.picpreview.ScaleHeight Then
         '     dstY = parent.picpreview.ScaleHeight
         ' End If
          

100       parent.picPreview.Line (SrcX, SrcY)-(dstx, dsty), vbYellow
          
          Const bombbounce = 5
          
          Dim i As Integer
110       For i = 1 To bombbounce
              Dim newval() As Double
120           newval = getNextIntersecion(val(0), val(1), val(2), val(3))
              
130           SrcX = dstx
140           SrcY = dsty
150           dstx = Int((newval(0) - lbx) * TILEW)
160           dsty = Int((newval(1) - lby) * TILEW)
170           parent.picPreview.Line (SrcX, SrcY)-(dstx, dsty), vbYellow
              
180           val = newval
190       Next
          
200       drawBombLine = False
End Sub

Function getNextIntersecion(startx As Double, startY As Double, vx As Double, vy As Double) As Double()
          Dim gni() As Double
10        ReDim gni(3)
          
          Dim curX As Double
          Dim curY As Double
          
20        curX = startx
30        curY = startY
          
          Dim collision As Boolean
          
40        Do While collision = False
50            DoEvents
60            curX = curX + vx
70            curY = curY + vy
          
80            If curX < 0 Then
90                collision = True

100               gni(0) = curX - vx
110               gni(1) = curY - vy
120               gni(2) = -vx
130               gni(3) = vy
                  
140               getNextIntersecion = gni
150               Exit Function
160           End If
              
170           If curX >= 1024 Then
180               collision = True
190               gni(0) = curX - vx
200               gni(1) = curY - vy
210               gni(2) = -vx
220               gni(3) = vy
                  
230               getNextIntersecion = gni
240               Exit Function
250           End If
              
260           If curY < 0 Then
270               collision = True
280               curY = 0
290               gni(0) = curX - vx
300               gni(1) = curY - vy
310               gni(2) = vx
320               gni(3) = -vy
                  
330               getNextIntersecion = gni
340               Exit Function
350               End If
              
360           If curY >= 1024 Then
370               collision = True
380               gni(0) = curX - vx
390               gni(1) = curY - vy
400               gni(2) = vx
410               gni(3) = -vy
                  
420               getNextIntersecion = gni
430               Exit Function
440           End If
              
450           If TileIsSolid(parent.getTile(Int(curX - vx), Int(curY))) Then
460               collision = True
470               gni(0) = curX - vx
480               gni(1) = curY - vy
490               gni(2) = vx
500               gni(3) = -vy
                  
510               getNextIntersecion = gni
520               Exit Function
530           End If
              
540           If TileIsSolid(parent.getTile(Int(curX), Int(curY - vy))) Then
550               collision = True
560               gni(0) = curX - vx
570               gni(1) = curY - vy
580               gni(2) = -vx
590               gni(3) = vy
                  
600               getNextIntersecion = gni
610               Exit Function
620           End If
              
630       Loop
End Function



Private Function RoundAimAngle() As Double
10        RoundAimAngle = Int(ship.aimangle / (PI / 20#)) * (PI / 20#)
End Function

Private Sub ChangeSpeed(up As Boolean, down As Boolean, bomb As Boolean)
          Dim curSpeed As Double
          Dim DeltaX As Double
          Dim DeltaY As Double
          
          Dim newSpeedX As Double
          Dim newSpeedY As Double

          Dim aimangle As Double
          
10        If up Or bomb Then
20            aimangle = RoundAimAngle
30        ElseIf down Then
40            aimangle = RoundAimAngle + PI
50        End If

60        If ship.turbo And (up Or down) Then
70            DeltaX = shipprops(ship.ship).MaximumThrust * Cos(aimangle)
80            DeltaY = shipprops(ship.ship).MaximumThrust * Sin(aimangle)
              
              'Drain energy from afterburners
90            ship.energy = ship.energy - shipprops(ship.ship).AfterBurnerEnergy
              
100       ElseIf up Or down Then
110           DeltaX = shipprops(ship.ship).InitialThrust * Cos(aimangle)
120           DeltaY = shipprops(ship.ship).InitialThrust * Sin(aimangle)
130       ElseIf bomb Then
140           DeltaX = -shipprops(ship.ship).BombThrust * Cos(aimangle)
150           DeltaY = -shipprops(ship.ship).BombThrust * Sin(aimangle)
160       Else
170           DeltaX = 0#
180           DeltaY = 0#
190       End If

200       newSpeedX = ship.vx + Fix(DeltaX)
210       newSpeedY = ship.vy + Fix(DeltaY)

220       curSpeed = Sqr(newSpeedX * newSpeedX + newSpeedY * newSpeedY)
          
          Dim MaxSpeedSqr As Double
          
230       If ship.turbo Then
240           MaxSpeedSqr = shipprops(ship.ship).MaximumSpeed
250       Else
260           MaxSpeedSqr = shipprops(ship.ship).InitialSpeed
270       End If

280       If curSpeed > MaxSpeedSqr Then
              'speed is now more than maxspeed, reduce speed, keeping same direction
290           newSpeedX = newSpeedX * (MaxSpeedSqr / curSpeed)
300           newSpeedY = newSpeedY * (MaxSpeedSqr / curSpeed)
310       End If

320       ship.vx = Fix(newSpeedX)
330       ship.vy = Fix(newSpeedY)
          
End Sub

Sub UpdateShipGraphics()
10        frmGeneral.picDefaultShips.AutoSize = True
20        frmGeneral.picDefaultShips.AutoSize = False
          
          Dim i As Integer
          Dim sizeX As Integer
          Dim sizeY As Integer
          
          Dim ret As String
30        ret = parent.lvz.SearchFile("ships", lvz_image)
          
          Dim srcPic As PictureBox
          
40        If ret <> "" Then
50            Call LoadPic(frmGeneral.pictemp, ret)
60            Set srcPic = frmGeneral.pictemp
70        Else
80            Set srcPic = frmGeneral.picDefaultShips
90        End If
          
100       For i = 0 To 7
110           frmGeneral.picShip(i).Cls
              
120           ret = parent.lvz.SearchFile("ship" & i + 1, lvz_image)
              
130           If ret <> "" Then
                  
140               frmGeneral.picShip(i).AutoSize = True
150               Call LoadPic(frmGeneral.picShip(i), ret)
160               frmGeneral.picShip(i).AutoSize = False
                  
170               sizeX = frmGeneral.picShip(i).width \ 10
180               sizeY = frmGeneral.picShip(i).height \ 4
                  
190           Else
                  'using default graphics for this ship
200               sizeX = srcPic.width \ 10
210               sizeY = srcPic.height \ 32
                  
220               frmGeneral.picShip(i).width = sizeX * 10
230               frmGeneral.picShip(i).height = sizeY * 4
                  
240               BitBlt frmGeneral.picShip(i).hDC, 0, 0, sizeX * 10, sizeY * 4, srcPic.hDC, 0, (sizeY * 4) * i, vbSrcCopy
250           End If
              
260           frmGeneral.picShip(i).Refresh
270           shipprops(i).Xsize = sizeX
280           shipprops(i).Ysize = sizeY
290       Next
          
300       Set srcPic = Nothing
End Sub

Sub ReadSettings()
          Dim i As Integer
          Dim ret As Integer
          
10        Call UpdateShipGraphics


20        For i = 0 To 7

30            shipprops(i).InitialThrust = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "InitialThrust", Default_InitialThrust)) * FPS / 10# 'nr of speed units gained per 10 ms
40            shipprops(i).MaximumThrust = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "MaximumThrust", Default_MaximumThrust)) * FPS / 10# 'afterburner thrust
              
50            shipprops(i).BombThrust = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "BombThrust")) * FPS / 10#
60            shipprops(i).BombFireDelay = parent.CFG.GetCfgSetting(ShipName(i), "BombFireDelay", 100) * 10 'setting is in 1/100th
              
              'a setting of 400 -> full rotation in 1 second
70            shipprops(i).Rotation = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "InitialRotation", Default_Rotation)) * PI * FPS / 200000#
              
              'a setting of 1 -> 1 pixel travelled per 10 seconds
80            shipprops(i).InitialSpeed = parent.CFG.GetCfgSetting(ShipName(i), "InitialSpeed", Default_InitialSpeed)
90            shipprops(i).MaximumSpeed = parent.CFG.GetCfgSetting(ShipName(i), "MaximumSpeed", Default_MaximumSpeed) 'afterburner max speed
              
100           shipprops(i).MaximumEnergy = parent.CFG.GetCfgSetting(ShipName(i), "InitialEnergy", Default_MaximumEnergy)
110           shipprops(i).Recharge = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "InitialRecharge", Default_Recharge)) * FPS / 10000#
              
120           shipprops(i).AfterBurnerEnergy = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "AfterburnerEnergy")) * FPS / 10000#
              
130           shipprops(i).BulletFireDelay = parent.CFG.GetCfgSetting(ShipName(i), "BulletFireDelay", Default_BulletFireDelay) * 10
140           shipprops(i).MultiFireDelay = parent.CFG.GetCfgSetting(ShipName(i), "MultiFireDelay", Default_MultiFireDelay) * 10
              
150           shipprops(i).BombFireEnergy = parent.CFG.GetCfgSetting(ShipName(i), "BombFireEnergy", Default_BombFireEnergy)
160           shipprops(i).BombFireEnergyUpgrade = parent.CFG.GetCfgSetting(ShipName(i), "BombFireEnergyUpgrade", Default_BombFireEnergyUpgrade)
              
170           shipprops(i).BulletFireEnergy = parent.CFG.GetCfgSetting(ShipName(i), "BulletFireEnergy", Default_BulletFireEnergy)
180           shipprops(i).MultiFireEnergy = parent.CFG.GetCfgSetting(ShipName(i), "MultiFireEnergy", Default_MultiFireEnergy)
              
190           shipprops(i).BombInitialLevel = parent.CFG.GetCfgSetting(ShipName(i), "InitialBombs", Default_BombInitialLevel)
200           shipprops(i).BombMaximumLevel = parent.CFG.GetCfgSetting(ShipName(i), "MaxBombs", Default_BombMaximumLevel)
              
210           shipprops(i).BulletInitialLevel = parent.CFG.GetCfgSetting(ShipName(i), "InitialGuns", Default_BulletInitialLevel)
220           shipprops(i).BulletMaximumLevel = parent.CFG.GetCfgSetting(ShipName(i), "MaxGuns", Default_BulletMaximumLevel)
          
              'collision radius of ship, convert the pixel setting to tiles
230           ret = CInt(parent.CFG.GetCfgSetting(ShipName(i), "Radius", 14))
240           If ret <= 0 Then
250               shipprops(i).radius = 14# / TILEW
260           Else
270               shipprops(i).radius = CDbl(ret) / TILEW
280           End If
290       Next
          
          'spawn settings
300       For i = 0 To 3
310           ret = CInt(parent.CFG.GetCfgSetting("Spawn", "Team" & CStr(i) & "-X"))
320           If ret <= 0 Then
330               mapprops.SpawnX(i) = 512
340           Else
350               mapprops.SpawnX(i) = ret
360           End If
              
370           ret = CInt(parent.CFG.GetCfgSetting("Spawn", "Team" & CStr(i) & "-Y"))
380           If ret <= 0 Then
390               mapprops.SpawnY(i) = 512
400           Else
410               mapprops.SpawnY(i) = ret
420           End If
              
430           ret = CInt(parent.CFG.GetCfgSetting("Spawn", "Team" & CStr(i) & "-radius"))
440           If ret <= 0 Then
450               mapprops.SpawnRadius(i) = CInt(parent.CFG.GetCfgSetting("Misc", "WarpRadiusLimit", Default_WarpRadiusLimit))
460           Else
470               mapprops.SpawnRadius(i) = ret
480           End If

490       Next
          
          'Map settings
500       mapprops.BounceFactor = 16# / CDbl(parent.CFG.GetCfgSetting("Misc", "BounceFactor", Default_BounceFactor))
              



End Sub

Private Function ShipName(shiptype As Integer) As String
10        Select Case shiptype
          Case 0
20            ShipName = "Warbird"
30        Case 1
40            ShipName = "Javelin"
50        Case 2
60            ShipName = "Spider"
70        Case 3
80            ShipName = "Leviathan"
90        Case 4
100           ShipName = "Terrier"
110       Case 5
120           ShipName = "Weasel"
130       Case 6
140           ShipName = "Lancaster"
150       Case 7
160           ShipName = "Shark"
170       Case Else
180           ShipName = "Spectator"
190       End Select
End Function

Sub FireBomb()
          Static lastBombFired As Long
          
          'Check the delay
10        If GetTickCount - lastBombFired >= shipprops(ship.ship).BombFireDelay Then
              'Check the energy needed
              Dim bombenergy As Long

20            bombenergy = shipprops(ship.ship).BombFireEnergy + (shipprops(ship.ship).BombFireEnergyUpgrade * (shipprops(ship.ship).BombInitialLevel - 1))
              
30            If ship.energy - bombenergy >= 0 Then
              
40                Call ChangeSpeed(False, False, True)
50                ship.energy = ship.energy - bombenergy
60                lastBombFired = GetTickCount
                  
70            End If
80        End If
End Sub

Sub FireBullet()
          Static lastBulletFired As Long
          'Check the delay
10        If GetTickCount - lastBulletFired >= shipprops(ship.ship).BulletFireDelay Then
              'Check the energy needed
              Dim bulletenergy As Long

20            bulletenergy = shipprops(ship.ship).BulletFireEnergy * shipprops(ship.ship).BulletInitialLevel
              
30            If ship.energy - bulletenergy >= 0 Then
              
40                ship.energy = ship.energy - bulletenergy
50                lastBulletFired = GetTickCount
                  
60            End If
70        End If
          
End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
20        Set undoch = Nothing
End Sub
