VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Switch/Replace Tiles"
   ClientHeight    =   3585
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRedraw 
      Caption         =   "Redraw walltiles"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox pictilesetlarge 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   0
      Top             =   240
      Width           =   1860
      Begin VB.Label lblswap 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   40
         X2              =   64
         Y1              =   16
         Y2              =   16
      End
      Begin VB.Line Line2 
         X1              =   56
         X2              =   64
         Y1              =   8
         Y2              =   16
      End
      Begin VB.Line Line3 
         X1              =   56
         X2              =   64
         Y1              =   24
         Y2              =   16
      End
      Begin VB.Line Line4 
         X1              =   48
         X2              =   40
         Y1              =   8
         Y2              =   16
      End
      Begin VB.Line Line5 
         X1              =   48
         X2              =   40
         Y1              =   24
         Y2              =   16
      End
      Begin VB.Shape shpright 
         BorderColor     =   &H0000FFFF&
         Height          =   510
         Left            =   1080
         Top             =   0
         Width           =   510
      End
      Begin VB.Shape shpleft 
         BorderColor     =   &H000000FF&
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
      Left            =   2160
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   9
      Top             =   120
      Width           =   4560
      Begin VB.Shape rightsel 
         BorderColor     =   &H0000FFFF&
         Height          =   240
         Left            =   240
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
   End
   Begin VB.OptionButton optreplaceleftright 
      Caption         =   "Replace left with right"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Switch or replace tiles"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox chkinselection 
      Caption         =   "In selection only"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton optswitchleftright 
      Caption         =   "Switch left  <-> right"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Label lblyellow 
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblred 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'defines which option to use
Enum replaceenum
    switchleftright
    replaceleftright
End Enum

'holds the left and right tile
Private parent As frmMain

Private tilesetleft As Integer
Private tilesetright As Integer

Public Sub setParent(Main As frmMain)
    Set parent = Main
End Sub



Private Sub cmdCancel_Click()
'Cancels the form
    Set parent = Nothing
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
'Proceed with replace or switching, and pass the attributes
'to the general form

    Dim using As replaceenum
    'holds the source and destination tile for replace
    '(when switching it doesn't matter which is src and
    'which is dest)
    Dim tilesrc As Integer
    Dim tiledest As Integer

    'no tiles bigger than 1 tile
    If TileIsSpecial(tilesetleft) Or TileIsSpecial(tilesetright) Then
        MessageBox "You can not use special tiles for tile switching/replacing", vbInformation
        Exit Sub
    End If

    'check for uselessness
    If tilesetleft = tilesetright And tilesetleft <= 256 Then
        MessageBox "You can't switch or replace a tile by the same tile... that's useless.", vbInformation
        Exit Sub
    End If

    'determine the operation
    If optswitchleftright.value Then
        using = switchleftright
    ElseIf optreplaceleftright.value Then
        using = replaceleftright
    End If

    'check to avoid filling the entire map with tiles
    If chkinselection.value = vbUnchecked Then
        If (using = replaceleftright And tilesetleft = 0) Or _
           (using = switchleftright And (tilesetleft = 0 Or tilesetright = 0)) Then
            MessageBox "You can't replace empty tiles all over the map.", vbInformation
            Exit Sub
        End If
    End If




    tilesrc = tilesetleft
    tiledest = tilesetright

    'pass the info to the executereplace method and execute it
    Call frmGeneral.ExecuteReplace(using, tilesrc, tiledest, chkinselection.value = vbChecked, chkRedraw.value = vbChecked)


End Sub

Private Sub Form_Activate()
'Updates the preview when activated
    DoEvents
    UpdatePreview
End Sub

Private Sub Form_Load()
'Copies the tileset from the general form
    Set Me.Icon = frmGeneral.Icon

    BitBlt pictileset.hDC, 0, 0, pictileset.width, pictileset.height, frmGeneral.cTileset.Pic_Tileset.hDC, 0, 0, vbSrcCopy
    pictileset.Refresh


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Cancel the form
    cmdCancel_Click
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
End Sub

Private Sub lblswap_Click()
    Dim tmp As Integer
    tmp = tilesetleft
    Call SetLeftSelection(tilesetright)
    Call SetRightSelection(tmp)
End Sub

Private Sub optreplaceleftright_Click()
'Hide one part of the arrow
    Line4.visible = False
    Line5.visible = False

    chkRedraw.Enabled = True

    UpdatePreview
End Sub

Private Sub optswitchleftright_Click()
'Show both sides of the arrow
    Line4.visible = True
    Line5.visible = True

    chkRedraw.Enabled = False

    UpdatePreview
End Sub

Private Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Selects the tile
    If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
        'out of range, don't select it
        Exit Sub
    End If

    'set the selected tile
    If Button = vbLeftButton Then
        SetLeftSelection ((Y \ TILEW) * 19 + (X \ TILEW) + 1)
    Else
        SetRightSelection ((Y \ TILEW) * 19 + (X \ TILEW) + 1)
    End If

End Sub

Private Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Do the same as mousedown
    If Button Then
        Call pictileset_MouseDown(Button, Shift, X, Y)
    End If

    Dim tilenr As Integer
    tilenr = (Y \ TILEW) * 19 + (X \ TILEW) + 1
    pictileset.tooltiptext = TilesetToolTipText(tilenr)

End Sub

Sub InitSelections()
    'TODO: IMPROVE
    If parent.tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        tilesetleft = parent.tileset.selection(vbLeftButton).tilenr
    ElseIf parent.tileset.selection(vbLeftButton).selectionType = TS_Walltiles Then
        tilesetleft = parent.tileset.selection(vbLeftButton).group + 259
    Else
        tilesetleft = 1
    End If
    
    If parent.tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        tilesetright = parent.tileset.selection(vbRightButton).tilenr
    ElseIf parent.tileset.selection(vbRightButton).selectionType = TS_Walltiles Then
        tilesetright = parent.tileset.selection(vbRightButton).group + 259
    Else
        tilesetright = 2
    End If

    'set the left and right tile
    Call SetLeftSelection(tilesetleft)
    Call SetRightSelection(tilesetright)
    
End Sub


Private Sub SetLeftSelection(tilenr)
'Set left selection of the tileset on the given tilenr
    If tilenr = 0 Then Exit Sub
    If tilenr > 256 And tilenr < 259 Then Exit Sub

    'move the left shape
    leftsel.visible = True
    leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
    leftsel.Top = ((tilenr - 1) \ 19) * TILEW

    If tilenr = 256 Then tilenr = 0
    tilesetleft = tilenr

    'if both shapes overlap make them white
    If tilesetleft = tilesetright Then
        rightsel.BorderColor = vbWhite
    Else
        rightsel.BorderColor = vbYellow
    End If

    'update the preview
    UpdatePreview

End Sub

Private Sub SetRightSelection(tilenr)
'Set right selection of the tileset on the given tilenr
    If tilenr = 0 Then Exit Sub
    If tilenr > 256 And tilenr < 259 Then Exit Sub

    'move the right shape
    rightsel.visible = True
    rightsel.Left = ((tilenr - 1) Mod 19) * TILEW
    rightsel.Top = ((tilenr - 1) \ 19) * TILEW

    If tilenr = 256 Then tilenr = 0
    tilesetright = tilenr

    'if both of them overlap make them white
    If tilesetleft = tilesetright Then
        rightsel.BorderColor = vbWhite
    Else
        rightsel.BorderColor = vbYellow
    End If

    UpdatePreview
End Sub

Private Sub UpdatePreview()
'Updates the preview
'stretch the 2 selected tiles from the tileset into the 2 large preview
    StretchBlt pictilesetlarge.hDC, shpleft.Left + 1, shpleft.Top + 1, shpleft.width - 2, shpleft.height - 2, pictileset.hDC, leftsel.Left, leftsel.Top, TILEW, TILEW, vbSrcCopy
    StretchBlt pictilesetlarge.hDC, shpright.Left + 1, shpright.Top + 1, shpright.width - 2, shpright.height - 2, pictileset.hDC, rightsel.Left, rightsel.Top, TILEW, TILEW, vbSrcCopy

    'update the numbers
    lblred.Caption = tilesetleft
    lblyellow.Caption = tilesetright
End Sub


