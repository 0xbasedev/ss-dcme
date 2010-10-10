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
10        If showing_details = True Then
20            Call Hide_Details
30        Else
40            Call Show_Details
50        End If
60        pictileset.Refresh
70        UpdatePreview
End Sub

Private Sub cmdOK_Click()
10        Unload Me
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon
End Sub

Sub setParent(c_parent As frmMain)
10        Set parent = c_parent
End Sub

Sub CountTiles()
'          Dim i As Integer

          'counters
          Dim total As Long
          Dim special As Long
          Dim m As Long
          Dim n As Long
          
          Dim inselection As Boolean
10        inselection = parent.sel.hasAlreadySelectedParts
      '    'reset the counters
      '    total = 0
      '    special = 0
      '    m = 0
      '    n = 0
          Dim lefttile As Integer
          Dim righttile As Integer

20        lefttile = parent.tileset.selection(vbLeftButton).tilenr
30        righttile = parent.tileset.selection(vbRightButton).tilenr
          
          'TODO: pass only a pointer to the array, return the total
40        ReDim tilecount(255)
50        total = parent.CountTiles(inselection, VarPtr(tilecount(0)))
          
          
      '    tilecount = parent.CountTiles(inSelection)
      '
      '    For i = 0 To 255
      '        total = total + tilecount(i)
      '    Next i

60        special = tilecount(TILE_WORMHOLE) + tilecount(TILE_LRG_ASTEROID) + tilecount(TILE_STATION)
70        m = tilecount(lefttile)
80        n = tilecount(righttile)

          'show the results
90        lblLeft.Caption = m
100       lblRight.Caption = n
110       lblTotal.Caption = total
120       lblSpecial.Caption = special
130       lblFlags.Caption = tilecount(TILE_FLAG)

140       lblInSelection.Visible = inselection

150       BitBlt pictileset.hDC, 0, 0, pictileset.Width, pictileset.Height, parent.pictileset.hDC, 0, 0, vbSrcCopy
160       Call SetLeftSelection(lefttile)
170       Call SetRightSelection(righttile)
180       Call SetCurrentSelection(1)

190       If tilecount(TILE_FLAG) > 256 Then
200           MessageBox "Warning: You have " & tilecount(TILE_FLAG) & " flags on your map. You will not be able to use your map if you have more than 256 flags.", vbOKOnly + vbExclamation, "Too many flags"
210       End If

220       Call Hide_Details
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Erase tilecount
          
20        Unload Me
End Sub

Sub SetLeftSelection(tilenr)
      'Set left selection of the tileset on the given tilenr
      'Eraser tile cannot be selected here

10        If tilenr = 0 Then Exit Sub
20        If tilenr >= 256 Then Exit Sub

          'move the left shape
30        leftsel.Visible = True
40        leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
50        leftsel.Top = ((tilenr - 1) \ 19) * TILEW

          'if both shapes overlap make them white
60        If ShapesOverlap(leftsel, rightsel) Then
70            rightsel.BorderColor = vbWhite
80        Else
90            rightsel.BorderColor = vbYellow
100       End If

End Sub

Sub SetRightSelection(tilenr)
      'Set right selection of the tileset on the given tilenr
      'Eraser tile cannot be selected here

10        If tilenr = 0 Then Exit Sub
20        If tilenr >= 256 Then Exit Sub

          'move the right shape
30        rightsel.Visible = True
40        rightsel.Left = ((tilenr - 1) Mod 19) * TILEW
50        rightsel.Top = ((tilenr - 1) \ 19) * TILEW

          'if both of them overlap make them white
60        If ShapesOverlap(leftsel, rightsel) Then
70            rightsel.BorderColor = vbWhite
80        Else
90            rightsel.BorderColor = vbYellow
100       End If

End Sub

Sub SetCurrentSelection(tilenr)
      'Set right selection of the tileset on the given tilenr
      'Eraser tile cannot be selected here

10        If tilenr = 0 Then Exit Sub
20        If tilenr >= 256 Then Exit Sub

          'move the blue shape
30        cursel.Visible = True
40        cursel.Left = ((tilenr - 1) Mod 19) * TILEW
50        cursel.Top = ((tilenr - 1) \ 19) * TILEW

60        tilesetcurrent = tilenr Mod 256

70        UpdatePreview

End Sub

Sub UpdatePreview()
      'Updates the preview
      'stretch the selected tile from the tileset into the large preview
10        StretchBlt pictilesetlarge.hDC, shpcur.Left + 1, shpcur.Top + 1, shpcur.Width - 2, shpcur.Height - 2, pictileset.hDC, cursel.Left, cursel.Top, TILEW, TILEW, vbSrcCopy

          'update the count
20        lblCount.Caption = tilecount(tilesetcurrent)
      '    If Len(TilesetToolTipText(tilesetcurrent)) > 36 Then
      '        lblTileNr.Caption = Mid(TilesetToolTipText(tilesetcurrent), 1, 36) & "..."
      '    Else
30            lblTileNr.Caption = TilesetToolTipText(tilesetcurrent)
              
      '    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
End Sub

Private Sub pictileset_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
      'Selects the tile
10        If Not (X > 0 And Y > 0 And X < pictileset.Width And Y < pictileset.Height) Then
              'out of range, don't select it
20            Exit Sub
30        End If

          'set the selected tile
40        If button = vbLeftButton Or button = vbRightButton Then
50            SetCurrentSelection ((Y \ TILEW) * 19 + ((X \ TILEW) + 1))
60        End If

End Sub

Private Sub pictileset_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
      'Do the same as mousedown
10        If button Then
20            Call pictileset_MouseDown(button, Shift, X, Y)
30        End If

          Dim tilenr As Integer
40        tilenr = (Y \ TILEW) * 19 + ((X \ TILEW) + 1)

50        pictileset.tooltiptext = TilesetToolTipText(tilenr)

End Sub

Private Sub Show_Details()
10        cmdDetails.Caption = "Hide Details"

20        frmCount.Width = 7590
30        frmCount.Height = 4095


40        showing_details = True

50        DoEvents
60        Call SetCurrentSelection(tilesetcurrent)
End Sub

Private Sub Hide_Details()
10        cmdDetails.Caption = "View Details"

20        frmCount.Width = 2835
30        frmCount.Height = 2750

40        showing_details = False
End Sub


