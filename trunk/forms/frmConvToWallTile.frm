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
10        Set parent = Main
End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub


Private Sub cmdOK_Click()
          Dim i As Integer

10        For i = optConvert.LBound To optConvert.UBound
20            If optConvert(i).value = True Then
30                Call SetSetting("LastConvert", CStr(i))
40                Call settings.SaveSettings
50                Call parent.sel.ConverttoWalltiles(curwall, i)
60            End If
70        Next

80        Unload Me
End Sub

Private Sub Form_Activate()
10        DoEvents
20        DrawWallTiles
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon

          Dim i As Integer
          Dim m As Integer

20        m = CInt(val(GetSetting("LastConvert", "0")))
30        For i = optConvert.LBound To optConvert.UBound
40            If i = m Then
50                optConvert(i).value = True
60            Else
70                optConvert(i).value = False
80            End If
90        Next

          'selwall.Visible = True
100       If parent.tileset.selection(vbLeftButton).selectionType = TS_Walltiles Then
110           curwall = parent.tileset.selection(vbLeftButton).group
120       ElseIf parent.tileset.selection(vbRightButton).selectionType = TS_Walltiles Then
130           curwall = parent.tileset.selection(vbRightButton).group
140       Else
150           curwall = 0
160       End If
170       If Not parent.walltiles.isValidSet(curwall) Then
180           For i = curwall To curwall + 7 Mod 8
190               If parent.walltiles.isValidSet(i) Then
200                   curwall = i
210                   GoTo skip
220               End If
230           Next
              'No valid walltiles set found
240           curwall = 0
250           selwall.visible = False
skip:
260       End If
270       selwall.Left = (curwall \ 4) * 64
280       selwall.Top = Int(curwall Mod 4) * 64

End Sub

Private Sub DrawWallTiles()
10        Call parent.walltiles.DrawWallTiles(picWalltiles.hDC, 4)
          
      '    Dim i As Integer
      '    For i = 0 To 7
      '        BitBlt picwalltiles.hdc, (i \ 4) * 4 * TILEW, (i Mod 4) * 4 * TILEW, 4 * TILEW, 4 * TILEW, parent.picwalltiles.hdc, i * 4 * TILEW, 0, vbSrcCopy
      '    Next
20        picWalltiles.Refresh
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
          
End Sub

Private Sub picWalltiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim tmp As Integer

10        If X < 0 Then X = 0
20        If X >= picWalltiles.width Then X = picWalltiles.width - 1
30        If Y < 0 Then Y = 0
40        If Y >= picWalltiles.height Then Y = picWalltiles.height - 1

50        selwall.visible = True

60        tmp = ((X \ (4 * TILEW)) * 4) + (Y \ 64)

70        If parent.walltiles.isValidSet(tmp) Then
              'New walltile selected
80            curwall = tmp
              'Move the selection rectangle
90            selwall.Left = (curwall \ 4) * 64
100           selwall.Top = Int(curwall Mod 4) * 64
110       End If


End Sub

Private Sub picWalltiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If Button Then Call picWalltiles_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picwalltiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If X >= selwall.Left And X <= selwall.Left + selwall.width And Y >= selwall.Top And Y <= selwall.Top + selwall.height Then
20            Call cmdOK_Click
30        End If
End Sub
