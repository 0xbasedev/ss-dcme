VERSION 5.00
Begin VB.Form frmCount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiles counted"
   ClientHeight    =   3885
   ClientLeft      =   180
   ClientTop       =   420
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pictilesetlarge 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   15
      Top             =   2760
      Width           =   660
      Begin VB.Shape shpcur 
         BorderColor     =   &H00FF0000&
         Height          =   510
         Left            =   0
         Top             =   0
         Width           =   510
      End
   End
   Begin VB.PictureBox pictileset 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   2880
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   17
      Top             =   120
      Width           =   4560
      Begin VB.Shape cursel 
         BorderColor     =   &H00FF0000&
         Height          =   240
         Left            =   480
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape leftsel 
         BorderColor     =   &H000000FF&
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape rightsel 
         BorderColor     =   &H0000FFFF&
         Height          =   240
         Left            =   240
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "View Details"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblFlags 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Flag tiles:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTileNr 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Count:"
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Selected Tile:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblInSelection 
      Alignment       =   2  'Center
      Caption         =   "Tiles counted in selection only."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblSpecial 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblRight 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Special tiles:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Total tiles:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Right tiles (yellow):"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Left tiles (red):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tilesetcurrent As Integer
Dim showing_details As Boolean

Dim tilecount() As Long
    
Dim parent As frmMain

Private Sub cmdDetails_Click()
    If showing_details = True Then
        Call Hide_Details
    Else
        Call Show_Details
    End If
    pictileset.Refresh
    UpdatePreview
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon
End Sub

Sub setParent(c_parent As frmMain)
    Set parent = c_parent
End Sub

Sub CountTiles()
'          Dim i As Integer

    'counters
    Dim total As Long
    Dim special As Long
    Dim m As Long
    Dim n As Long
    
    Dim inselection As Boolean
    inselection = parent.sel.hasAlreadySelectedParts
'    'reset the counters
'    total = 0
'    special = 0
'    m = 0
'    n = 0
    Dim lefttile As Integer
    Dim righttile As Integer

    lefttile = parent.tileset.selection(vbLeftButton).tilenr
    righttile = parent.tileset.selection(vbRightButton).tilenr
    
    'TODO: pass only a pointer to the array, return the total
    ReDim tilecount(255)
    total = parent.CountTiles(inselection, VarPtr(tilecount(0)))
    
    
'    tilecount = parent.CountTiles(inSelection)
'
'    For i = 0 To 255
'        total = total + tilecount(i)
'    Next i

    special = tilecount(TILE_WORMHOLE) + tilecount(TILE_LRG_ASTEROID) + tilecount(TILE_STATION)
    m = tilecount(lefttile)
    n = tilecount(righttile)

    'show the results
    lblLeft.Caption = m
    lblRight.Caption = n
    lblTotal.Caption = total
    lblSpecial.Caption = special
    lblFlags.Caption = tilecount(TILE_FLAG)

    lblInSelection.visible = inselection

    BitBlt pictileset.hDC, 0, 0, pictileset.width, pictileset.height, parent.pictileset.hDC, 0, 0, vbSrcCopy
    Call SetLeftSelection(lefttile)
    Call SetRightSelection(righttile)
    Call SetCurrentSelection(1)

    If tilecount(TILE_FLAG) > 256 Then
        MessageBox "Warning: You have " & tilecount(TILE_FLAG) & " flags on your map. You will not be able to use your map if you have more than 256 flags.", vbOKOnly + vbExclamation, "Too many flags"
    End If

    Call Hide_Details
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Erase tilecount
    
    Unload Me
End Sub

Sub SetLeftSelection(tilenr)
'Set left selection of the tileset on the given tilenr
'Eraser tile cannot be selected here

    If tilenr = 0 Then Exit Sub
    If tilenr >= 256 Then Exit Sub

    'move the left shape
    leftsel.visible = True
    leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
    leftsel.Top = ((tilenr - 1) \ 19) * TILEW

    'if both shapes overlap make them white
    If ShapesOverlap(leftsel, rightsel) Then
        rightsel.BorderColor = vbWhite
    Else
        rightsel.BorderColor = vbYellow
    End If

End Sub

Sub SetRightSelection(tilenr)
'Set right selection of the tileset on the given tilenr
'Eraser tile cannot be selected here

    If tilenr = 0 Then Exit Sub
    If tilenr >= 256 Then Exit Sub

    'move the right shape
    rightsel.visible = True
    rightsel.Left = ((tilenr - 1) Mod 19) * TILEW
    rightsel.Top = ((tilenr - 1) \ 19) * TILEW

    'if both of them overlap make them white
    If ShapesOverlap(leftsel, rightsel) Then
        rightsel.BorderColor = vbWhite
    Else
        rightsel.BorderColor = vbYellow
    End If

End Sub

Sub SetCurrentSelection(tilenr)
'Set right selection of the tileset on the given tilenr
'Eraser tile cannot be selected here

    If tilenr = 0 Then Exit Sub
    If tilenr >= 256 Then Exit Sub

    'move the blue shape
    cursel.visible = True
    cursel.Left = ((tilenr - 1) Mod 19) * TILEW
    cursel.Top = ((tilenr - 1) \ 19) * TILEW

    tilesetcurrent = tilenr Mod 256

    UpdatePreview

End Sub

Sub UpdatePreview()
'Updates the preview
'stretch the selected tile from the tileset into the large preview
    StretchBlt pictilesetlarge.hDC, shpcur.Left + 1, shpcur.Top + 1, shpcur.width - 2, shpcur.height - 2, pictileset.hDC, cursel.Left, cursel.Top, TILEW, TILEW, vbSrcCopy

    'update the count
    lblCount.Caption = tilecount(tilesetcurrent)
'    If Len(TilesetToolTipText(tilesetcurrent)) > 36 Then
'        lblTileNr.Caption = Mid(TilesetToolTipText(tilesetcurrent), 1, 36) & "..."
'    Else
        lblTileNr.Caption = TilesetToolTipText(tilesetcurrent)
        
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
End Sub

Private Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Selects the tile
    If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
        'out of range, don't select it
        Exit Sub
    End If

    'set the selected tile
    If Button = vbLeftButton Or Button = vbRightButton Then
        SetCurrentSelection ((Y \ TILEW) * 19 + ((X \ TILEW) + 1))
    End If

End Sub

Private Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Do the same as mousedown
    If Button Then
        Call pictileset_MouseDown(Button, Shift, X, Y)
    End If

    Dim tilenr As Integer
    tilenr = (Y \ TILEW) * 19 + ((X \ TILEW) + 1)

    pictileset.tooltiptext = TilesetToolTipText(tilenr)

End Sub

Private Sub Show_Details()
    cmdDetails.Caption = "Hide Details"

    frmCount.width = 7590
    frmCount.height = 4095


    showing_details = True

    DoEvents
    Call SetCurrentSelection(tilesetcurrent)
End Sub

Private Sub Hide_Details()
    cmdDetails.Caption = "View Details"

    frmCount.width = 2835
    frmCount.height = 2750

    showing_details = False
End Sub


