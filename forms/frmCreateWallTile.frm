VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateWallTile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Walltiles"
   ClientHeight    =   5025
   ClientLeft      =   135
   ClientTop       =   405
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pictilesetlarge 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1680
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   4
      Top             =   3960
      Width           =   1740
      Begin VB.Label lblswap 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.Shape shpleft 
         BorderColor     =   &H000000FF&
         Height          =   510
         Left            =   0
         Top             =   0
         Width           =   510
      End
      Begin VB.Shape shpright 
         BorderColor     =   &H0000FFFF&
         Height          =   510
         Left            =   1080
         Top             =   0
         Width           =   510
      End
      Begin VB.Line Line5 
         X1              =   48
         X2              =   40
         Y1              =   24
         Y2              =   16
      End
      Begin VB.Line Line4 
         X1              =   48
         X2              =   40
         Y1              =   8
         Y2              =   16
      End
      Begin VB.Line Line3 
         X1              =   56
         X2              =   64
         Y1              =   24
         Y2              =   16
      End
      Begin VB.Line Line2 
         X1              =   56
         X2              =   64
         Y1              =   8
         Y2              =   16
      End
      Begin VB.Line Line1 
         X1              =   40
         X2              =   64
         Y1              =   16
         Y2              =   16
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Walltiles"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Walltiles"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Current"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.PictureBox piccurrent 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   480
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   3840
      Width           =   960
      Begin VB.Shape shpwall2 
         BorderColor     =   &H0000FFFF&
         Height          =   240
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpwall 
         BorderColor     =   &H00FF0000&
         Height          =   240
         Left            =   240
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox picwalltiles 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   4920
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   120
      Width           =   1920
      Begin VB.Shape selwall 
         BorderColor     =   &H000000FF&
         Height          =   960
         Left            =   0
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Frame frmCurrent 
      Caption         =   "Current"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox pictileset 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   240
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   0
      Top             =   120
      Width           =   4560
      Begin VB.Shape shpcur 
         BorderColor     =   &H00FF0000&
         Height          =   240
         Left            =   360
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape rightsel 
         BorderColor     =   &H0000FFFF&
         Height          =   240
         Left            =   480
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape leftsel 
         BorderColor     =   &H000000FF&
         Height          =   240
         Left            =   0
         Top             =   360
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5640
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmCreateWallTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Current map
Public curmap As frmMain

'Current wall set index
Dim curwall As Integer

'Store walltiles data temporarly
Public tempwalltiles As walltiles

Dim tilesetleft As Integer
Dim tilesetright As Integer

Dim oldtilesetX As Integer
Dim oldtilesetY As Integer

'size of the tiles selected
Dim multTileLeftx As Integer
Dim multTileLefty As Integer
Dim multTileRightx As Integer
Dim multTileRighty As Integer

Dim tempwalltileschanged As Boolean


Sub setParent(Main As frmMain)
    Set curmap = Main
    Set tempwalltiles = New walltiles
    Call tempwalltiles.setParent(Main)
End Sub

Private Sub cmdCancel_Click()
    If tempwalltileschanged Then
        If MessageBox("Are you sure you want to cancel changes made to your walltiles?", vbYesNo + vbInformation, "Cancel changes") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call ApplyWalltiles
    Unload Me
End Sub


Private Sub DrawWallTiles()
    Call tempwalltiles.DrawWallTiles(picWalltiles.hDC, 2)
    
'    Dim i As Integer
'    For i = 0 To 7
'        'For each wall set, first draw the default image
'        BitBlt picwalltiles.hDC, (i \ 4) * 4 * TILEW, (i Mod 4) * 4 * TILEW, TILEW * 4, TILEW * 4, frmGeneral.picdefault.hDC, 0, 0, vbSrcCopy
'
'        Dim a As Integer
'        For a = 0 To 15
'            Dim val As Integer
'
'            'Get corresponding wall tile
'            val = tempwalltiles.getWallTile(i, tempwalltiles.tileconvert(a))
'
'            'If a wall tile is set to that spot, draw it
'            If val <> 0 Then
'                BitBlt picwalltiles.hDC, (((i \ 4) * 4) + (a Mod 4)) * TILEW, (((a \ 4) + i * 4) - (i \ 4) * TILEW) * TILEW, TILEW, TILEW, pictileset.hDC, ((val - 1) Mod 19) * TILEW, ((val - 1) \ 19) * TILEW, vbSrcCopy
'            End If
'        Next
'    Next
    picWalltiles.Refresh

    'Posistion the rectangle on the selected wall set
    selwall.Left = (curwall Mod 2) * WT_SETW
    selwall.Top = (curwall \ 2) * WT_SETH
    
    Call tempwalltiles.DrawWallTilesSet(piccurrent.hDC, 0, 0, curwall)
    'Draw the current wall set on piccurrent
'    BitBlt piccurrent.hDC, 0, 0, 4 * TILEW, 4 * TILEW, picwalltiles.hDC, (curwall \ 4) * 64, Int(curwall Mod 4) * 64, vbSrcCopy
    piccurrent.Refresh
End Sub


Private Sub cmdClear_Click()
'Clear current wall set
    If MessageBox("Are you sure you want to clear the current wall tile?", vbYesNo, "Clear wall tile") = vbYes Then
        tempwalltiles.clearWallTile (curwall)
        Call DrawWallTiles
    End If
End Sub

Private Sub cmdClearAll_Click()
'Clear all wall sets
    If MessageBox("Are you sure you want to clear all wall tiles?", vbYesNo, "Clear wall tiles") = vbYes Then
        Dim i As Integer
        For i = 0 To 7
            tempwalltiles.clearWallTile (i)
        Next
        Call DrawWallTiles
    End If
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo errh

    cd.DialogTitle = "Load walltiles..."
    cd.flags = cdlOFNHideReadOnly
    cd.filename = curmap.Caption & "_walls.wtl"
    cd.Filter = "*.wtl|*.wtl"

    cd.ShowOpen

    If cd.filename <> "" Then
        Call tempwalltiles.LoadWallTiles(cd.filename)
        tempwalltileschanged = True
    End If

    DrawWallTiles

    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errh

    cd.DialogTitle = "Save walltiles..."
    cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    cd.filename = curmap.Caption & ".wtl"
    cd.Filter = "*.wtl|*.wtl"

    cd.ShowSave

    If cd.filename <> "" Then
        Call tempwalltiles.SaveWallTiles(cd.filename)
    End If

    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
    End If
End Sub

Private Sub Form_Activate()
    'Copy tileset
    If Not curmap Is Nothing Then
        BitBlt pictileset.hDC, 0, 0, pictileset.Width, pictileset.Height, curmap.pictileset.hDC, 0, 0, vbSrcCopy
    End If
    
    pictileset.Refresh

    DoEvents
    'Update preview
    Call DrawWallTiles

    UpdatePreview
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon



    tempwalltileschanged = False

End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'If mouse is outside piccurrent (i.e. on the form), hide the rectangles
    shpwall.Visible = False
    shpwall2.Visible = False
    shpcur.Visible = False
End Sub


Private Sub ApplyWalltiles()
    Set curmap.walltiles = tempwalltiles
    Call curmap.DrawWallTilesPreview
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set tempwalltiles = Nothing
    Set curmap = Nothing
End Sub

Private Sub frmCurrent_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'If mouse is outside piccurrent (i.e. on the form), hide the rectangles
    shpwall.Visible = False
    shpwall2.Visible = False
    shpcur.Visible = False
End Sub


Private Sub lblswap_Click()
    Dim tmp As Integer
    Dim tmpsizex As Integer
    Dim tmpsizey As Integer
    tmp = tilesetleft
    tmpsizex = multTileLeftx
    tmpsizey = multTileLefty
    Call SetLeftSelection(tilesetright, multTileRightx, multTileRighty)
    Call SetRightSelection(tmp, tmpsizex, tmpsizey)
End Sub

Private Sub piccurrent_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim assignto As Integer
    Dim tile As Integer
    Dim multTileX As Integer
    Dim multTileY As Integer

    Dim i As Integer
    Dim j As Integer

    'Set the tile and size X / Y to the left or right values
    If button = vbLeftButton Then
        tile = tilesetleft
        multTileX = multTileLeftx
        multTileY = multTileLefty
    ElseIf button = vbRightButton Then
        tile = tilesetright
        multTileX = multTileRightx
        multTileY = multTileRighty
    Else
        Exit Sub
    End If





    'If the selection is larger than 4 in width or height, abort
    If (X \ TILEW) + multTileX > 4 Then Exit Sub
    If (Y \ TILEW) + multTileY > 4 Then Exit Sub




    'Calculate on which tile we clicked
    '0....3
    '......
    '12..15
    assignto = ((Y \ TILEW) * 4 + (X \ TILEW))

    Dim tmp As Integer

    'Assign the tile to the corresponding walltile index
    For j = 0 To multTileY - 1
        For i = 0 To multTileX - 1
            Call tempwalltiles.setWallTile(curwall, tempwalltiles.tileconvert(assignto + i + (4 * j)), tile + i + (19 * j))
        Next
    Next
    Call DrawWallTiles
    tempwalltileschanged = True

End Sub


Private Sub piccurrent_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'Only allow tile 0 to be used on mousemove, to quickly clear multiple wall tiles
    If button = vbLeftButton Then
        If tilesetleft = 0 Then Call piccurrent_MouseDown(button, Shift, X, Y)
    ElseIf button = vbRightButton Then
        If tilesetright = 0 Then Call piccurrent_MouseDown(button, Shift, X, Y)
    End If

    'Hide the rectangle if tiles selection cannot fit in walltiles preview
    If (X \ TILEW) + multTileLeftx > 4 Or (Y \ TILEW) + multTileLefty > 4 Then
        shpwall.Visible = False
    Else
        shpwall.Visible = True
    End If

    'Place and resize the main rectangle
    shpwall.Left = (X \ TILEW) * TILEW
    shpwall.Top = (Y \ TILEW) * TILEW
    shpwall.Width = TILEW * multTileLeftx
    shpwall.Height = TILEW * multTileLefty

    If multTileLeftx = 0 Then multTileLeftx = 1
    If multTileRightx = 0 Then multTileRightx = 1
    If multTileLefty = 0 Then multTileLefty = 1
    If multTileRighty = 0 Then multTileRighty = 1

    If multTileLeftx <> multTileRightx Or multTileLefty <> multTileRighty Then
        'If the size of Left selection and Right selection are different,
        'Show the second rectangle (yellow), place it, and make the first one red
        If (X \ TILEW) + multTileRightx > 4 Or (Y \ TILEW) + multTileRighty > 4 Then
            shpwall2.Visible = False
        Else
            shpwall2.Visible = True
        End If
        shpwall2.Left = (X \ TILEW) * TILEW + 1
        shpwall2.Top = (Y \ TILEW) * TILEW + 1
        shpwall2.Width = TILEW * multTileRightx - 2
        shpwall2.Height = TILEW * multTileRighty - 2
        shpwall.BorderColor = vbRed
    Else
        'Both selections are the same, hide the yellow rectangle and keep the blue rectangle
        shpwall2.Visible = False
        shpwall.BorderColor = vbBlue
    End If


    'Show selected tile on the tileset
    Dim tmptilenr As Integer
    Dim assignto As Integer

    assignto = ((Y \ TILEW) * 4 + (X \ TILEW))
    tmptilenr = tempwalltiles.getWallTile(curwall, tempwalltiles.tileconvert(assignto))
    If tmptilenr <> 0 Then
        shpcur.Left = ((tmptilenr - 1) Mod 19) * TILEW
        shpcur.Top = ((tmptilenr - 1) \ 19) * TILEW
        shpcur.Visible = True
    Else
        shpcur.Visible = False
    End If

End Sub

Private Sub pictileset_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (X > 0 And Y > 0 And X < pictileset.Width And Y < pictileset.Height) Then
        'not in the boundaries of the picture
        Exit Sub
    End If


    'indicate the mouse is down

    oldtilesetX = (X \ TILEW) + 1
    oldtilesetY = Y \ TILEW

    'set the selected tile
    If button = vbLeftButton Then
        Call SetLeftSelection(oldtilesetY * 19 + oldtilesetX, 1, 1)
    Else
        Call SetRightSelection(oldtilesetY * 19 + oldtilesetX, 1, 1)
    End If
End Sub

Private Sub pictileset_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (X > 0 And Y > 0 And X < pictileset.Width And Y < pictileset.Height) Then
        'not in the boundaries of the picture
        Exit Sub
    End If


    Dim tilenr As Integer
    tilenr = (Y \ TILEW) * 19 + ((X \ TILEW) + 1)
    pictileset.tooltiptext = TilesetToolTipText(tilenr)



    If button Then

        Dim sizeX As Integer
        Dim sizeY As Integer

        Dim curtilesetX As Integer
        Dim curtilesetY As Integer
        curtilesetX = (X \ TILEW) + 1
        curtilesetY = Y \ TILEW

        sizeX = Abs(curtilesetX - oldtilesetX)
        sizeY = Abs(curtilesetY - oldtilesetY)

        If curtilesetX > oldtilesetX Then
            curtilesetX = oldtilesetX
        End If
        If curtilesetY > oldtilesetY Then
            curtilesetY = oldtilesetY
        End If
        If curtilesetX < oldtilesetX And sizeX >= 3 Then
            curtilesetX = oldtilesetX - 3
        End If
        If curtilesetY < oldtilesetY And sizeY >= 3 Then
            curtilesetY = oldtilesetY - 3
        End If

        If sizeX > 3 Then sizeX = 3
        If sizeY > 3 Then sizeY = 3

        Dim holdtileset As Boolean
        holdtileset = False
        If (curtilesetX <= 8 And 8 <= curtilesetX + sizeX) And _
           (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
            holdtileset = True
        End If
        If (curtilesetX <= 10 And 10 <= curtilesetX + sizeX) And _
           (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
            holdtileset = True
        End If
        If (curtilesetX <= 9 And 9 <= curtilesetX + sizeX) And _
           (curtilesetY <= 13 And 13 <= curtilesetY + sizeY) Then
            holdtileset = True
        End If
        If (curtilesetX <= 11 And 11 <= curtilesetX + sizeX) And _
           (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
            holdtileset = True
        End If
        If tilenr > 256 Then
            holdtileset = True
        End If
        If holdtileset Then
            sizeX = 0
            sizeY = 0
            curtilesetX = oldtilesetX
            curtilesetY = oldtilesetY
        End If




        'set the selected tile
        If button = vbLeftButton Then
            Call SetLeftSelection(curtilesetY * 19 + curtilesetX, sizeX + 1, sizeY + 1)
        Else
            Call SetRightSelection(curtilesetY * 19 + curtilesetX, sizeX + 1, sizeY + 1)
        End If

    End If
End Sub



Sub SetLeftSelection(tilenr As Integer, sizeX As Integer, sizeY As Integer)
'Set left selection of the tileset on the given tilenr
    If tilenr = 0 Then Exit Sub
    If tilenr > 256 Then Exit Sub
    If tilenr = 217 Or tilenr = 219 Or tilenr = 220 Then Exit Sub

    'move the left shape
    leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
    leftsel.Top = ((tilenr - 1) \ 19) * TILEW

    leftsel.Width = TILEW * sizeX
    leftsel.Height = TILEW * sizeY

    tilesetleft = tilenr Mod 256
    multTileLeftx = sizeX
    multTileLefty = sizeY


    'if the left and right tiles are the same, show a white shape to
    'indicate they are the same
    If tilesetleft = tilesetright And multTileLeftx = multTileRightx And multTileLefty = multTileRighty Then
        rightsel.BorderColor = vbWhite
    Else
        rightsel.BorderColor = vbYellow
    End If

    UpdatePreview
End Sub

Sub SetRightSelection(tilenr As Integer, sizeX As Integer, sizeY As Integer)
'Set right selection of the tileset on the given tilenr
    If tilenr = 0 Then Exit Sub
    If tilenr > 256 Then Exit Sub
    If tilenr = 217 Or tilenr = 219 Or tilenr = 220 Then Exit Sub

    'move the right shape
    rightsel.Left = ((tilenr - 1) Mod 19) * TILEW
    rightsel.Top = ((tilenr - 1) \ 19) * TILEW

    rightsel.Width = TILEW * sizeX
    rightsel.Height = TILEW * sizeY

    tilesetright = tilenr Mod 256
    multTileRightx = sizeX
    multTileRighty = sizeY

    'if both of them overlap make them white
    If tilesetleft = tilesetright And multTileLeftx = multTileRightx And multTileLefty = multTileRighty Then
        rightsel.BorderColor = vbWhite
    Else
        rightsel.BorderColor = vbYellow
    End If

    UpdatePreview
End Sub





Private Sub picWalltiles_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'New walltile selected
    curwall = X \ 64 + (Y \ 64) * 2
    'Move the selection rectangle
    selwall.Left = (curwall \ 4) * 64
    selwall.Top = Int(curwall Mod 4) * 64
    'Update preview
    Call DrawWallTiles
End Sub


Sub Init(newcurwall As Integer) ', tileleft As Integer, tileright As Integer)
'Define on which walltile to set focus at first
    curwall = newcurwall

    'Set left and right tiles to the same as on the map
'    tilesetleft = tileleft
'    tilesetright = tileright
'
'    'Make sure special tiles are not selected
'    If tilesetleft = 217 Or tilesetleft = 219 Or tilesetleft = 220 Then tilesetleft = 1
'    If tilesetright = 217 Or tilesetright = 219 Or tilesetright = 220 Then tilesetright = 2
    tilesetleft = 1
    tilesetright = 2
    
    Call SetLeftSelection(tilesetleft, 1, 1)
    Call SetRightSelection(tilesetright, 1, 1)
    Call DrawWallTiles
End Sub




Private Sub UpdatePreview()
'Updates the preview
'stretch the 2 selected tiles from the tileset into the 2 large preview
    StretchBlt pictilesetlarge.hDC, shpleft.Left + 1, shpleft.Top + 1, shpleft.Width - 2, shpleft.Height - 2, pictileset.hDC, leftsel.Left, leftsel.Top, TILEW, TILEW, vbSrcCopy
    StretchBlt pictilesetlarge.hDC, shpright.Left + 1, shpright.Top + 1, shpright.Width - 2, shpright.Height - 2, pictileset.hDC, rightsel.Left, rightsel.Top, TILEW, TILEW, vbSrcCopy
End Sub

