VERSION 5.00
Begin VB.Form frmConvToWallTile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convert to walltiles"
   ClientHeight    =   3495
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   278
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optConvert 
      Caption         =   "Complete Walltiles"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "This option will redraw the selected walltiles set within the selection"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Convert"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton optConvert 
      Caption         =   "Only Current Walltiles"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "This option will only convert tiles that are part of one of the walltiles set"
      Top             =   360
      Width           =   1815
   End
   Begin VB.OptionButton optConvert 
      Caption         =   "All Tiles"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "This option will convert all tiles within the selection"
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox picwalltiles 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   240
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      ToolTipText     =   "Click on a valid walltiles set to proceed"
      Top             =   960
      Width           =   3840
      Begin VB.Shape selwall 
         BorderColor     =   &H000000FF&
         Height          =   960
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmConvToWallTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim parent As frmMain
Dim curwall As Integer

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim i As Integer

    For i = optConvert.LBound To optConvert.UBound
        If optConvert(i).value = True Then
            Call SetSetting("LastConvert", CStr(i))
            Call settings.SaveSettings
            Call parent.sel.ConverttoWalltiles(curwall, i)
        End If
    Next

    Unload Me
End Sub

Private Sub Form_Activate()
    DoEvents
    DrawWallTiles
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    Dim i As Integer
    Dim m As Integer

    m = CInt(val(GetSetting("LastConvert", "0")))
    For i = optConvert.LBound To optConvert.UBound
        If i = m Then
            optConvert(i).value = True
        Else
            optConvert(i).value = False
        End If
    Next

    'selwall.Visible = True
    If parent.tileset.selection(vbLeftButton).selectionType = TS_Walltiles Then
        curwall = parent.tileset.selection(vbLeftButton).group
    ElseIf parent.tileset.selection(vbRightButton).selectionType = TS_Walltiles Then
        curwall = parent.tileset.selection(vbRightButton).group
    Else
        curwall = 0
    End If
    If Not parent.walltiles.isValidSet(curwall) Then
        For i = curwall To curwall + 7 Mod 8
            If parent.walltiles.isValidSet(i) Then
                curwall = i
                GoTo skip
            End If
        Next
        'No valid walltiles set found
        curwall = 0
        selwall.visible = False
skip:
    End If
    selwall.Left = (curwall \ 4) * 64
    selwall.Top = Int(curwall Mod 4) * 64

End Sub

Private Sub DrawWallTiles()
    Call parent.walltiles.DrawWallTiles(picWalltiles.hDC, 4)
    
'    Dim i As Integer
'    For i = 0 To 7
'        BitBlt picwalltiles.hdc, (i \ 4) * 4 * TILEW, (i Mod 4) * 4 * TILEW, 4 * TILEW, 4 * TILEW, parent.picwalltiles.hdc, i * 4 * TILEW, 0, vbSrcCopy
'    Next
    picWalltiles.Refresh
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
    
End Sub

Private Sub picWalltiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Integer

    If X < 0 Then X = 0
    If X >= picWalltiles.width Then X = picWalltiles.width - 1
    If Y < 0 Then Y = 0
    If Y >= picWalltiles.height Then Y = picWalltiles.height - 1

    selwall.visible = True

    tmp = ((X \ (4 * TILEW)) * 4) + (Y \ 64)

    If parent.walltiles.isValidSet(tmp) Then
        'New walltile selected
        curwall = tmp
        'Move the selection rectangle
        selwall.Left = (curwall \ 4) * 64
        selwall.Top = Int(curwall Mod 4) * 64
    End If


End Sub

Private Sub picWalltiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then Call picWalltiles_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picwalltiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= selwall.Left And X <= selwall.Left + selwall.width And Y >= selwall.Top And Y <= selwall.Top + selwall.height Then
        Call cmdOK_Click
    End If
End Sub
