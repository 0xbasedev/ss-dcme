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
    isRunning = Not c_stop
End Function

Sub Class_Initialize()
    
    c_stop = True
    
    'init position to -1 so we can know we need to warp for the first time in StartRun
    ship.X = -1
    ship.Y = -1
    
    ship.vx = 0
    ship.vy = 0
    
    ship.ship = CInt(GetSetting("TestMapShip", "1")) - 1
    ship.freq = 0
End Sub

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub setShipType(i As Integer)
    ship.ship = i
    ship.energy = shipprops(ship.ship).MaximumEnergy
    frmGeneral.optShip(i).value = True
End Sub

Sub StartRun()
   'main loop
    Dim tick As Long
    Dim sleepval As Long
    
    sleepval = FPS

    Call ReadSettings
    
    If ship.X = -1 Or ship.Y = -1 Then
        Call WarpShip
    End If
    

    'Add the width of a tile to the pictureboxes so it can be pixel accurate.
    'The updateLevel function will need to work on one more tile for smooth display

    
    With parent.picPreview
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
    End With
    
    parent.shpcursor.visible = False
    
    'used for debugging, real-time vs execution time
    Dim starttick As Long
    'used to calculate frame per second
    Dim displayedframes As Integer
    Dim cyclescount As Long
    
    'set zoom to normal
    If parent.magnifier.zoom <> 1 Then
        Call parent.magnifier.SetZoom(1#, False)
    End If
    
    GetKeys
    
    c_stop = False
    
    ship.energy = shipprops(ship.ship).MaximumEnergy
    
    parent.UpdateLevel
    frmGeneral.UpdateToolBarButtons
    
    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y
    
    Do While Not c_stop
        tick = GetTickCount
        While GetTickCount - tick < sleepval
            'Debug.Print "waiting"
            Sleep 1
            DoEvents
        Wend
        
        frmGeneral.Label6.visible = True
        frmGeneral.Label7.visible = True
        
        frmGeneral.lblposition.Caption = Format(ship.X, "0.0") & " - " & Format(ship.Y, "0.0")
        frmGeneral.Label7.Caption = " Speed: " & ship.vx & " ; " & ship.vy & " (" & Format(Sqr(ship.vx * ship.vx + ship.vy * ship.vy), "0.00") & ") Aim:" & RoundAimAngle * 180# / PI & Chr(176) & " Energy:" & Int(ship.energy)
        DoEvents
        
        
        GetKeys
        
        If c_stop Then Exit Sub
        
        CollisionCheck
        
        UpdateShipPos
        
        Call setMapPosFromShipPos
        
        CheckShipWarp

        If ship.energy < shipprops(ship.ship).MaximumEnergy Then
            If ship.energy + shipprops(ship.ship).Recharge >= shipprops(ship.ship).MaximumEnergy Then
                ship.energy = shipprops(ship.ship).MaximumEnergy
            Else
                ship.energy = ship.energy + shipprops(ship.ship).Recharge
            End If
        End If
        
        If sleepval > 0 Then
            Call parent.UpdateLevel(True, False)

            Call parent.UpdatePreview(False, True)


            lbx = parent.Hscr.value \ parent.currenttilew
            hbx = (parent.Hscr.value + parent.picPreview.width) \ parent.currenttilew
            lby = parent.Vscr.value \ parent.currenttilew
            hby = (parent.Vscr.value + parent.picPreview.height) \ parent.currenttilew
            
        '    frmGeneral.Label6.Caption = ship.Y
            Dim offsetX As Integer
            Dim offsetY As Integer
            
            If Int(ship.X) = lbx + (hbx - lbx) \ 2 And _
               Int(ship.Y) = lby + (hby - lby) \ 2 Then
                offsetX = -(Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
                offsetY = -(Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)
            ElseIf Int(ship.X) = lbx + (hbx - lbx) \ 2 Then
                offsetX = -(Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
                offsetY = 0
            ElseIf Int(ship.Y) = lby + (hby - lby) \ 2 Then
                offsetX = 0
                offsetY = -(Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)
            Else
                offsetX = 0
                offsetY = 0
            End If
            
            Call parent.UpdatePreview(False, False)
'            Call BitBlt(parent.picPreview.hDC, offsetX, offsetY, parent.piclevel.width + parent.currenttilew, parent.piclevel.height + parent.currenttilew, parent.piclevel.hDC, 0, 0, vbSrcCopy)
            
'            If parent.ShowLVZ Then
'                Call parent.UpdateLVZPreview(lbx, lby, hbx, hby, True)
'            End If
'
            Call drawShip
            
            If drawBombLine Then
                'Dim t As Long
                't = GetTickCount
                Call drawBombPath
                
            End If
            
            'to calculate FPS properly
            displayedframes = displayedframes + 1
            
            parent.picPreview.Refresh

        ElseIf sleepval < -FPSTIMEOUT Then
            'something is wrong, 5 seconds are missing... just forget them (happens often in the IDE,
            'when pausing execution.. could also happen if the application is temporarly stopped for some reason
            sleepval = 0
        Else
            'we got to skip some frames
        End If
        
        'calculate FPS every second
        'FPS  -  real time / desired time
        If cyclescount * FPS >= 1000 Then
            frmGeneral.Label6.Caption = "FPS: " & displayedframes & " - " & GetTickCount - starttick & "ms / " & cyclescount * FPS & "ms"
            starttick = GetTickCount
            displayedframes = 0
            cyclescount = 0
        End If
        
        cyclescount = cyclescount + 1
        
        sleepval = FPS - (GetTickCount - tick) + sleepval
    Loop
    
    'call getkeys again in case user is still holding Q or W, it will end the drawing and append to undo if needed
    GetKeys
    
    Call parent.UpdatePreview
   
End Sub

Sub StopRun()
    c_stop = True
    
    If Not bDEBUG Then
        frmGeneral.Label6.visible = False
        frmGeneral.Label7.visible = False
    End If

    If isDrawing Then
        'we were drawing tiles before, append to undo
        Call parent.undoredo.AddToUndo(undoch, UNDO_SHIPDRAW)
        isDrawing = False
    End If
    
    'update preview picture
    parent.picPreview.width = (parent.ScaleWidth - parent.Vscr.width)
    parent.picPreview.height = (parent.ScaleHeight - parent.Hscr.height)

    'update level picture according to preview picture
    With parent.picPreview
'        parent.piclevel.width = .width
'        parent.piclevel.height = .height
'        parent.piclvzunder.width = .width
'        parent.piclvzunder.height = .height
'        parent.piclvzover.width = .width
'        parent.piclvzover.height = .height
    End With
        

    
    parent.UpdateLevel
    frmGeneral.UpdateToolBarButtons
        
End Sub

Sub GetKeys()

    If Not isRunning Then
        If isDrawing Then
            'we were drawing tiles before, append to undo
            Call parent.undoredo.AddToUndo(undoch, UNDO_SHIPDRAW)
        End If
        GetAsyncKeyState vbKeyQ
        GetAsyncKeyState vbKeyW
        isDrawing = False
    Else
    
        If GetAsyncKeyState(vbKeyShift) <> 0 And (GetAsyncKeyState(vbKeyUp) <> 0 Or GetAsyncKeyState(vbKeyDown) <> 0) Then
            ship.turbo = True
        Else
            ship.turbo = False
        End If
        
        If GetAsyncKeyState(vbKeyUp) <> 0 Then
            KeyPress (vbKeyUp)
        ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
            KeyPress (vbKeyDown)
        Else
            Call ChangeSpeed(False, False, False) 'cancel any remaining afterburner
        End If
        If GetAsyncKeyState(vbKeyLeft) <> 0 Then
            KeyPress (vbKeyLeft)
        End If
        If GetAsyncKeyState(vbKeyRight) <> 0 Then
            KeyPress (vbKeyRight)
        End If
        If GetAsyncKeyState(vbKeyEscape) <> 0 Then
            KeyPress (vbKeyEscape)
        End If
        If GetAsyncKeyState(vbKeyInsert) <> 0 Then
            KeyPress (vbKeyInsert)
        End If
        If GetAsyncKeyState(vbKeyB) <> 0 Then
            KeyPress (vbKeyB)
        End If
        If GetAsyncKeyState(vbKeyTab) <> 0 Then
            KeyPress (vbKeyTab)
        End If
        If GetAsyncKeyState(vbKeyControl) <> 0 Then
            KeyPress (vbKeyControl)
        End If
        
        If GetAsyncKeyState(vbKeyQ) <> 0 Then
            If Not isDrawing Then
                Set undoch = New Changes
                parent.undoredo.ResetRedo
                isDrawing = True
            End If
            KeyPress (vbKeyQ)
        ElseIf GetAsyncKeyState(vbKeyW) <> 0 Then
            If Not isDrawing Then
                Set undoch = New Changes
                parent.undoredo.ResetRedo
                isDrawing = True
            End If
            KeyPress (vbKeyW)
        Else
            'neither Q or W are pressed
            If isDrawing Then
                'we were drawing tiles before, append to undo
                Call parent.undoredo.AddToUndo(undoch, UNDO_SHIPDRAW)
            End If

            isDrawing = False
        End If
        
        Dim i As Integer
        For i = 0 To 7
            If GetAsyncKeyState(vbKey1 + i) <> 0 Then
                KeyPress (vbKey1 + i)
                Exit For
            End If
        Next
        
    End If
    
End Sub


Sub KeyPress(KeyAscii As Integer)
    If Not isRunning Then Exit Sub
    
    drawBombLine = False
    
    Select Case KeyAscii
        Case vbKeyUp

            Call ChangeSpeed(True, False, False)
        
        Case vbKeyDown
            
            Call ChangeSpeed(False, True, False)
        
        Case vbKeyRight
            ship.aimangle = ship.aimangle + shipprops(ship.ship).Rotation
            If ship.aimangle > 2 * PI Then ship.aimangle = ship.aimangle - 2 * PI
            
        Case vbKeyLeft
            ship.aimangle = ship.aimangle - shipprops(ship.ship).Rotation
            If ship.aimangle < 0 Then ship.aimangle = ship.aimangle + 2 * PI
        
        Case vbKeyInsert
            If ship.energy = shipprops(ship.ship).MaximumEnergy Then
                Call WarpShip
            End If
        Case vbKeyEscape
            StopRun
        
        Case vbKeyB
            drawBombLine = True
            
        Case vbKeyQ
            If parent.tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
                Call parent.setTile(Int(ship.X), Int(ship.Y), parent.tileset.selection(vbLeftButton).tilenr, undoch, True)
                Call parent.UpdateLevelTile(Int(ship.X), Int(ship.Y), False, True)
            End If
        Case vbKeyW
            If parent.tileset.selection(vbRightButton).selectionType = TS_Tiles Then
                Call parent.setTile(Int(ship.X), Int(ship.Y), parent.tileset.selection(vbRightButton).tilenr, undoch, True)
                Call parent.UpdateLevelTile(Int(ship.X), Int(ship.Y), False, True)
            End If
        Case vbKey1 To vbKey8
            If ship.energy = shipprops(ship.ship).MaximumEnergy Then
                Call setShipType(KeyAscii - vbKey1)
            End If
        
        Case vbKeyTab
            FireBomb
        Case vbKeyControl
            FireBullet
            
    End Select
End Sub

Private Sub CollisionCheck()
    On Error GoTo errh:
    
    If Not isRunning Then Exit Sub
    
    'calculate tile left of ship
    Dim testTileX As Integer
    Dim testTileY As Integer
    Dim i As Integer
    Dim j As Integer
    
    'LEFT CHECK
    '-----------
    If Int(ship.X + ship.vx * FPS / (10000# * TILEW)) < Int(ship.X) Then
        'we're about to hit the wall
            For i = 0 To (shipprops(ship.ship).Ysize \ TILEW) - 1
                testTileY = Round(ship.Y + (((shipprops(ship.ship).Ysize Mod TILEW) / 2) / TILEW)) + i 'Int(ship.X + (((shipprops(ship.ship).Xsize Mod TileW) / 2) / TileW)) + i
                testTileX = Int(ship.X) - 1

                'check indices first
                If testTileX < 0 Then
                    Call BounceX
                    GoTo continue1
                ElseIf testTileY < 0 Then
                    Call BounceY
                Else
                    'we have a valid x
                    If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
                        Call BounceX
                        GoTo continue1
                    End If
                End If
            Next
    End If
    
continue1:
    'TOP CHECK
    '------------
    If Int(ship.Y + ship.vy * FPS / (10000# * TILEW)) < Int(ship.Y) Then
        'we're about to hit the wall
            For i = 0 To (shipprops(ship.ship).Xsize \ TILEW) - 1
                testTileX = Round(ship.X + (((shipprops(ship.ship).Xsize Mod TILEW) / 2) / TILEW)) + i 'Int(ship.X + (((shipprops(ship.ship).Xsize Mod TileW) / 2) / TileW)) + i
                testTileY = Int(ship.Y) - 1

                'check indices first
                If testTileY < 0 Then
                    Call BounceY
                    GoTo continue2
                Else
                    'we have a valid y
                    If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
                        ship.Y = testTileY + 1.05
                        Call BounceY
                        GoTo continue2
                    End If
                End If
            Next
    End If

continue2:
    'RIGHT CHECK
    '-----------
    If Int(ship.X + ship.vx * FPS / (10000# * TILEW)) > Int(ship.X) Then
        'we're about to hit the wall
            For i = 0 To (shipprops(ship.ship).Ysize \ TILEW) - 1
                testTileX = Round(ship.X + (shipprops(ship.ship).Xsize / TILEW))
                testTileY = Int(ship.Y + (((shipprops(ship.ship).Ysize Mod TILEW) / 2) / TILEW)) + i
            
                'check indices first
                If testTileX > 1023 Then
                    Call BounceX
                    GoTo continue3
                Else
                    'we have a valid x
                    If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
                        Call BounceX
                        GoTo continue3
                    End If
                End If
            Next
    End If

continue3:
    'BOTTOM CHECK
    '------------
    If Int(ship.Y + ship.vy * FPS / (10000# * TILEW)) > Int(ship.Y) Then
        'we're about to hit the wall
            For i = 0 To (shipprops(ship.ship).Xsize \ TILEW) - 1
                testTileX = Round(ship.X + (((shipprops(ship.ship).Xsize Mod TILEW) / 2) / TILEW)) + i
                testTileY = Int(ship.Y + (shipprops(ship.ship).Ysize \ TILEW) - (((shipprops(ship.ship).Ysize Mod TILEW) / 2) / TILEW)) + 1
            
                'check indices first
                If testTileY > 1023 Then
                    Call BounceY
                    GoTo continue4
                Else
                    'we have a valid y
                    If frmGeneral.chkTileCollision.value = checked And TileIsSolid(parent.getTile(testTileX, testTileY)) Then
                        Call BounceY
                        GoTo continue4
                    End If
                End If
            Next
    End If
    
continue4:

Exit Sub
errh:

End Sub

Private Sub BounceX()
    If Not isRunning Then Exit Sub
    
    ship.vx = -ship.vx * mapprops.BounceFactor
    ship.vy = ship.vy * mapprops.BounceFactor
End Sub

Private Sub BounceY()
    If Not isRunning Then Exit Sub
    
    ship.vy = -ship.vy * mapprops.BounceFactor
    ship.vx = ship.vx * mapprops.BounceFactor
End Sub

Private Sub UpdateShipPos()
    If Not isRunning Then Exit Sub
    
    ship.X = ship.X + ship.vx * FPS / (10000# * TILEW)
    ship.Y = ship.Y + ship.vy * FPS / (10000# * TILEW)
End Sub

Private Sub setMapPosFromShipPos()
    On Error GoTo errh
    
    If Not isRunning Then Exit Sub
    
    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y
    lbx = parent.Hscr.value
    hbx = (parent.Hscr.value + parent.picPreview.width)
    lby = parent.Vscr.value
    hby = (parent.Vscr.value + parent.picPreview.height)
        
    Dim xpos As Single
    Dim ypos As Single
    xpos = ship.X - ((hbx - lbx) / parent.currenttilew) / 2
    ypos = ship.Y - ((hby - lby) / parent.currenttilew) / 2
    
'90        If xpos < 0 Then xpos = 0
'100       If ypos < 0 Then ypos = 0
'110       If xpos + (hbx - lbx) / parent.currenttilew > 1023 Then xpos = 1023 - (hbx - lbx - 1)
'120       If ypos + (hby - lby) / parent.currenttilew > 1023 Then ypos = 1023 - (hby - lby - 1)
    

    Call parent.SetScrollbarValues(xpos * parent.currenttilew, ypos * parent.currenttilew, False)

    
    Exit Sub
errh:
    
End Sub

Private Sub drawShip()
    If Not isRunning Then Exit Sub
    
    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y
    lbx = parent.Hscr.value
    hbx = (parent.Hscr.value + parent.picPreview.width)
    lby = parent.Vscr.value
    hby = (parent.Vscr.value + parent.picPreview.height)

    Dim xSrc As Integer
    Dim ySrc As Integer
    Dim shpAngD As Integer
    shpAngD = 90 + ((RoundAimAngle / PI) * 180)
    shpAngD = shpAngD Mod 360
    
    xSrc = Int(((shpAngD Mod 90) / 9) * (shipprops(ship.ship).Xsize) / shipprops(ship.ship).Xsize) * shipprops(ship.ship).Xsize
    ySrc = Int((shpAngD / 360) * 4) * shipprops(ship.ship).Ysize ' + _
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
    
    If frmGeneral.chkTileCollision.value = checked Then
        For i = Int(ship.X - shipprops(ship.ship).radius) To RoundAway(ship.X + shipprops(ship.ship).radius)
            For j = Int(ship.Y - shipprops(ship.ship).radius) To RoundAway(ship.Y + shipprops(ship.ship).radius)
                If i >= 0 And i <= 1023 And j >= 0 And j <= 1023 Then
                    If TileIsWarp(parent.getTile(i, j)) Then
                        Call WarpShip
                        Exit Sub
                    End If
                End If
            Next
        Next
    End If
End Sub


Sub WarpShip(Optional warpX As Integer = -1, Optional warpY As Integer = -1)
    Dim toX As Integer
    Dim toY As Integer
    Dim tries As Long
    tries = 1
    
    Randomize
    Do
        If warpX = -1 Then
            toX = CInt(Rnd() * (mapprops.SpawnRadius(ship.freq Mod 4) * 2 + 1)) + mapprops.SpawnX(ship.freq Mod 4)
        Else
            toX = warpX
        End If
        If warpY < 0 Or warpY > 1023 Then
            toY = CInt(Rnd() * (mapprops.SpawnRadius(ship.freq Mod 4) * 2 + 1)) + mapprops.SpawnY(ship.freq Mod 4)
        Else
            toY = warpY
        End If
        
        If toX < 0 Then toX = 0
        If toX > 1023 Then toX = 1023
        If toY < 0 Then toY = 0
        If toY > 1023 Then toY = 1023
        
        tries = tries + 1

    'try coordinates until tile is free, abort after 10000 tries
    Loop While (TileIsSolid(parent.getTile(toX, toY)) And tries < 10000)
    '<<<<<<<<<<<<<<<<<<<<<<<<
    'TODO: SHIP MUST HAVE ENOUGH SPACE AROUND IT, NOT ONLY AT THE TILE
    
    ship.energy = 0
    
    ship.X = toX
    ship.Y = toY
    ship.vx = 0
    ship.vy = 0
    
    'randomize orientation
    ship.aimangle = 2 * PI * Rnd()

End Sub


Private Sub drawBombPath()
    Dim lbx As Integer    'lowestval for x
   ' Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
   ' Dim hby As Integer    'highestval for y
    lbx = parent.Hscr.value / parent.currenttilew
   ' hbx = (parent.Hscr.value + parent.picpreview.Width) \ parent.currenttilew
    lby = parent.Vscr.value / parent.currenttilew
   ' hby = (parent.Vscr.value + parent.picpreview.Height) \ parent.currenttilew


    Dim val() As Double
    
    Dim vx As Double
    Dim vy As Double
    vy = Sin(RoundAimAngle)
    vx = Cos(RoundAimAngle)
    
    val = getNextIntersecion(Int(ship.X) + (shipprops(ship.ship).Xsize / TILEW) / 2, Int(ship.Y) + (shipprops(ship.ship).Ysize / TILEW) / 2, vx, vy) '2 * (ship.vx / shipprops(ship.ship).MaxSpeed), 2 * (ship.vy / shipprops(ship.ship).MaxSpeed))
    
    Dim SrcX As Integer
    Dim SrcY As Integer
    Dim dstx As Integer
    Dim dsty As Integer
    
    SrcX = Int(ship.X - lbx) * TILEW + (shipprops(ship.ship).Xsize / 2) '+ (Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
    SrcY = Int(ship.Y - lby) * TILEW + (shipprops(ship.ship).Ysize / 2) '+ (Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)

    
    dstx = ((val(0) - lbx) * TILEW) - (Int(ship.X * parent.currenttilew) Mod parent.currenttilew)
    dsty = ((val(1) - lby) * TILEW) - (Int(ship.Y * parent.currenttilew) Mod parent.currenttilew)
    
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
    

    parent.picPreview.Line (SrcX, SrcY)-(dstx, dsty), vbYellow
    
    Const bombbounce = 5
    
    Dim i As Integer
    For i = 1 To bombbounce
        Dim newval() As Double
        newval = getNextIntersecion(val(0), val(1), val(2), val(3))
        
        SrcX = dstx
        SrcY = dsty
        dstx = Int((newval(0) - lbx) * TILEW)
        dsty = Int((newval(1) - lby) * TILEW)
        parent.picPreview.Line (SrcX, SrcY)-(dstx, dsty), vbYellow
        
        val = newval
    Next
    
    drawBombLine = False
End Sub

Function getNextIntersecion(startx As Double, startY As Double, vx As Double, vy As Double) As Double()
    Dim gni() As Double
    ReDim gni(3)
    
    Dim curX As Double
    Dim curY As Double
    
    curX = startx
    curY = startY
    
    Dim collision As Boolean
    
    Do While collision = False
        DoEvents
        curX = curX + vx
        curY = curY + vy
    
        If curX < 0 Then
            collision = True

            gni(0) = curX - vx
            gni(1) = curY - vy
            gni(2) = -vx
            gni(3) = vy
            
            getNextIntersecion = gni
            Exit Function
        End If
        
        If curX >= 1024 Then
            collision = True
            gni(0) = curX - vx
            gni(1) = curY - vy
            gni(2) = -vx
            gni(3) = vy
            
            getNextIntersecion = gni
            Exit Function
        End If
        
        If curY < 0 Then
            collision = True
            curY = 0
            gni(0) = curX - vx
            gni(1) = curY - vy
            gni(2) = vx
            gni(3) = -vy
            
            getNextIntersecion = gni
            Exit Function
            End If
        
        If curY >= 1024 Then
            collision = True
            gni(0) = curX - vx
            gni(1) = curY - vy
            gni(2) = vx
            gni(3) = -vy
            
            getNextIntersecion = gni
            Exit Function
        End If
        
        If TileIsSolid(parent.getTile(Int(curX - vx), Int(curY))) Then
            collision = True
            gni(0) = curX - vx
            gni(1) = curY - vy
            gni(2) = vx
            gni(3) = -vy
            
            getNextIntersecion = gni
            Exit Function
        End If
        
        If TileIsSolid(parent.getTile(Int(curX), Int(curY - vy))) Then
            collision = True
            gni(0) = curX - vx
            gni(1) = curY - vy
            gni(2) = -vx
            gni(3) = vy
            
            getNextIntersecion = gni
            Exit Function
        End If
        
    Loop
End Function



Private Function RoundAimAngle() As Double
    RoundAimAngle = Int(ship.aimangle / (PI / 20#)) * (PI / 20#)
End Function

Private Sub ChangeSpeed(up As Boolean, down As Boolean, bomb As Boolean)
    Dim curSpeed As Double
    Dim DeltaX As Double
    Dim DeltaY As Double
    
    Dim newSpeedX As Double
    Dim newSpeedY As Double

    Dim aimangle As Double
    
    If up Or bomb Then
        aimangle = RoundAimAngle
    ElseIf down Then
        aimangle = RoundAimAngle + PI
    End If

    If ship.turbo And (up Or down) Then
        DeltaX = shipprops(ship.ship).MaximumThrust * Cos(aimangle)
        DeltaY = shipprops(ship.ship).MaximumThrust * Sin(aimangle)
        
        'Drain energy from afterburners
        ship.energy = ship.energy - shipprops(ship.ship).AfterBurnerEnergy
        
    ElseIf up Or down Then
        DeltaX = shipprops(ship.ship).InitialThrust * Cos(aimangle)
        DeltaY = shipprops(ship.ship).InitialThrust * Sin(aimangle)
    ElseIf bomb Then
        DeltaX = -shipprops(ship.ship).BombThrust * Cos(aimangle)
        DeltaY = -shipprops(ship.ship).BombThrust * Sin(aimangle)
    Else
        DeltaX = 0#
        DeltaY = 0#
    End If

    newSpeedX = ship.vx + Fix(DeltaX)
    newSpeedY = ship.vy + Fix(DeltaY)

    curSpeed = Sqr(newSpeedX * newSpeedX + newSpeedY * newSpeedY)
    
    Dim MaxSpeedSqr As Double
    
    If ship.turbo Then
        MaxSpeedSqr = shipprops(ship.ship).MaximumSpeed
    Else
        MaxSpeedSqr = shipprops(ship.ship).InitialSpeed
    End If

    If curSpeed > MaxSpeedSqr Then
        'speed is now more than maxspeed, reduce speed, keeping same direction
        newSpeedX = newSpeedX * (MaxSpeedSqr / curSpeed)
        newSpeedY = newSpeedY * (MaxSpeedSqr / curSpeed)
    End If

    ship.vx = Fix(newSpeedX)
    ship.vy = Fix(newSpeedY)
    
End Sub

Sub UpdateShipGraphics()
    frmGeneral.picDefaultShips.AutoSize = True
    frmGeneral.picDefaultShips.AutoSize = False
    
    Dim i As Integer
    Dim sizeX As Integer
    Dim sizeY As Integer
    
    Dim ret As String
    ret = parent.lvz.SearchFile("ships", lvz_image)
    
    Dim srcPic As PictureBox
    
    If ret <> "" Then
        Call LoadPic(frmGeneral.pictemp, ret)
        Set srcPic = frmGeneral.pictemp
    Else
        Set srcPic = frmGeneral.picDefaultShips
    End If
    
    For i = 0 To 7
        frmGeneral.picShip(i).Cls
        
        ret = parent.lvz.SearchFile("ship" & i + 1, lvz_image)
        
        If ret <> "" Then
            
            frmGeneral.picShip(i).AutoSize = True
            Call LoadPic(frmGeneral.picShip(i), ret)
            frmGeneral.picShip(i).AutoSize = False
            
            sizeX = frmGeneral.picShip(i).width \ 10
            sizeY = frmGeneral.picShip(i).height \ 4
            
        Else
            'using default graphics for this ship
            sizeX = srcPic.width \ 10
            sizeY = srcPic.height \ 32
            
            frmGeneral.picShip(i).width = sizeX * 10
            frmGeneral.picShip(i).height = sizeY * 4
            
            BitBlt frmGeneral.picShip(i).hDC, 0, 0, sizeX * 10, sizeY * 4, srcPic.hDC, 0, (sizeY * 4) * i, vbSrcCopy
        End If
        
        frmGeneral.picShip(i).Refresh
        shipprops(i).Xsize = sizeX
        shipprops(i).Ysize = sizeY
    Next
    
    Set srcPic = Nothing
End Sub

Sub ReadSettings()
    Dim i As Integer
    Dim ret As Integer
    
    Call UpdateShipGraphics


    For i = 0 To 7

        shipprops(i).InitialThrust = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "InitialThrust", Default_InitialThrust)) * FPS / 10# 'nr of speed units gained per 10 ms
        shipprops(i).MaximumThrust = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "MaximumThrust", Default_MaximumThrust)) * FPS / 10# 'afterburner thrust
        
        shipprops(i).BombThrust = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "BombThrust")) * FPS / 10#
        shipprops(i).BombFireDelay = parent.CFG.GetCfgSetting(ShipName(i), "BombFireDelay", 100) * 10 'setting is in 1/100th
        
        'a setting of 400 -> full rotation in 1 second
        shipprops(i).Rotation = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "InitialRotation", Default_Rotation)) * PI * FPS / 200000#
        
        'a setting of 1 -> 1 pixel travelled per 10 seconds
        shipprops(i).InitialSpeed = parent.CFG.GetCfgSetting(ShipName(i), "InitialSpeed", Default_InitialSpeed)
        shipprops(i).MaximumSpeed = parent.CFG.GetCfgSetting(ShipName(i), "MaximumSpeed", Default_MaximumSpeed) 'afterburner max speed
        
        shipprops(i).MaximumEnergy = parent.CFG.GetCfgSetting(ShipName(i), "InitialEnergy", Default_MaximumEnergy)
        shipprops(i).Recharge = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "InitialRecharge", Default_Recharge)) * FPS / 10000#
        
        shipprops(i).AfterBurnerEnergy = CDbl(parent.CFG.GetCfgSetting(ShipName(i), "AfterburnerEnergy")) * FPS / 10000#
        
        shipprops(i).BulletFireDelay = parent.CFG.GetCfgSetting(ShipName(i), "BulletFireDelay", Default_BulletFireDelay) * 10
        shipprops(i).MultiFireDelay = parent.CFG.GetCfgSetting(ShipName(i), "MultiFireDelay", Default_MultiFireDelay) * 10
        
        shipprops(i).BombFireEnergy = parent.CFG.GetCfgSetting(ShipName(i), "BombFireEnergy", Default_BombFireEnergy)
        shipprops(i).BombFireEnergyUpgrade = parent.CFG.GetCfgSetting(ShipName(i), "BombFireEnergyUpgrade", Default_BombFireEnergyUpgrade)
        
        shipprops(i).BulletFireEnergy = parent.CFG.GetCfgSetting(ShipName(i), "BulletFireEnergy", Default_BulletFireEnergy)
        shipprops(i).MultiFireEnergy = parent.CFG.GetCfgSetting(ShipName(i), "MultiFireEnergy", Default_MultiFireEnergy)
        
        shipprops(i).BombInitialLevel = parent.CFG.GetCfgSetting(ShipName(i), "InitialBombs", Default_BombInitialLevel)
        shipprops(i).BombMaximumLevel = parent.CFG.GetCfgSetting(ShipName(i), "MaxBombs", Default_BombMaximumLevel)
        
        shipprops(i).BulletInitialLevel = parent.CFG.GetCfgSetting(ShipName(i), "InitialGuns", Default_BulletInitialLevel)
        shipprops(i).BulletMaximumLevel = parent.CFG.GetCfgSetting(ShipName(i), "MaxGuns", Default_BulletMaximumLevel)
    
        'collision radius of ship, convert the pixel setting to tiles
        ret = CInt(parent.CFG.GetCfgSetting(ShipName(i), "Radius", 14))
        If ret <= 0 Then
            shipprops(i).radius = 14# / TILEW
        Else
            shipprops(i).radius = CDbl(ret) / TILEW
        End If
    Next
    
    'spawn settings
    For i = 0 To 3
        ret = CInt(parent.CFG.GetCfgSetting("Spawn", "Team" & CStr(i) & "-X"))
        If ret <= 0 Then
            mapprops.SpawnX(i) = 512
        Else
            mapprops.SpawnX(i) = ret
        End If
        
        ret = CInt(parent.CFG.GetCfgSetting("Spawn", "Team" & CStr(i) & "-Y"))
        If ret <= 0 Then
            mapprops.SpawnY(i) = 512
        Else
            mapprops.SpawnY(i) = ret
        End If
        
        ret = CInt(parent.CFG.GetCfgSetting("Spawn", "Team" & CStr(i) & "-radius"))
        If ret <= 0 Then
            mapprops.SpawnRadius(i) = CInt(parent.CFG.GetCfgSetting("Misc", "WarpRadiusLimit", Default_WarpRadiusLimit))
        Else
            mapprops.SpawnRadius(i) = ret
        End If

    Next
    
    'Map settings
    mapprops.BounceFactor = 16# / CDbl(parent.CFG.GetCfgSetting("Misc", "BounceFactor", Default_BounceFactor))
        



End Sub

Private Function ShipName(shiptype As Integer) As String
    Select Case shiptype
    Case 0
        ShipName = "Warbird"
    Case 1
        ShipName = "Javelin"
    Case 2
        ShipName = "Spider"
    Case 3
        ShipName = "Leviathan"
    Case 4
        ShipName = "Terrier"
    Case 5
        ShipName = "Weasel"
    Case 6
        ShipName = "Lancaster"
    Case 7
        ShipName = "Shark"
    Case Else
        ShipName = "Spectator"
    End Select
End Function

Sub FireBomb()
    Static lastBombFired As Long
    
    'Check the delay
    If GetTickCount - lastBombFired >= shipprops(ship.ship).BombFireDelay Then
        'Check the energy needed
        Dim bombenergy As Long

        bombenergy = shipprops(ship.ship).BombFireEnergy + (shipprops(ship.ship).BombFireEnergyUpgrade * (shipprops(ship.ship).BombInitialLevel - 1))
        
        If ship.energy - bombenergy >= 0 Then
        
            Call ChangeSpeed(False, False, True)
            ship.energy = ship.energy - bombenergy
            lastBombFired = GetTickCount
            
        End If
    End If
End Sub

Sub FireBullet()
    Static lastBulletFired As Long
    'Check the delay
    If GetTickCount - lastBulletFired >= shipprops(ship.ship).BulletFireDelay Then
        'Check the energy needed
        Dim bulletenergy As Long

        bulletenergy = shipprops(ship.ship).BulletFireEnergy * shipprops(ship.ship).BulletInitialLevel
        
        If ship.energy - bulletenergy >= 0 Then
        
            ship.energy = ship.energy - bulletenergy
            lastBulletFired = GetTickCount
            
        End If
    End If
    
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
