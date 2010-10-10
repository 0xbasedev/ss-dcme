VERSION 5.00
Begin VB.Form frmResize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resize Selection"
   ClientHeight    =   4650
   ClientLeft      =   495
   ClientTop       =   780
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReceiveFocus 
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Text            =   "This is the textbox that receives focus when form is loaded (else it will mess up the focus booleans of the other textboxes)"
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkmaintainratio 
      Caption         =   "Lock aspect ratio"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   4200
      Width           =   1335
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Tile size"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Percentage of original"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Current size"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label lbloriginalheight 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbloriginalwidth 
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ComboBox cmbResizeType 
      Height          =   315
      ItemData        =   "frmResize.frx":0000
      Left            =   1800
      List            =   "frmResize.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3690
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   4575
      Begin VB.TextBox txtPercentWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Top             =   335
         Width           =   975
      End
      Begin VB.TextBox txtPercentHeight 
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   335
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "x         Height:"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4575
      Begin VB.TextBox txtTileHeight 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   335
         Width           =   975
      End
      Begin VB.TextBox txtTileWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   335
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "x         Height:"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Resize type:"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Dim percHfocus As Boolean
Dim percWfocus As Boolean
Dim tileHfocus As Boolean
Dim tileWfocus As Boolean

Sub setParent(map As frmMain)
10        Set parent = map
End Sub

Private Sub Form_Load()

10        Set Me.Icon = frmGeneral.Icon

          Dim oSize As Integer
20        oSize = CInt(GetSetting("OptSize", "0"))
30        optSize(oSize).value = True

40        lbloriginalwidth.Caption = parent.sel.getBoundaries.Right - parent.sel.getBoundaries.Left
50        lbloriginalheight.Caption = parent.sel.getBoundaries.Bottom - parent.sel.getBoundaries.Top

60        If optSize(0).value Then
70            txtTileWidth.Text = GetSetting("TileWidth", lbloriginalwidth.Caption)
80            txtTileHeight.Text = GetSetting("TileHeight", lbloriginalheight.Caption)

90            txtPercentWidth.Text = (val(txtTileWidth.Text) / val(lbloriginalwidth.Caption)) * 100
100           txtPercentHeight.Text = (val(txtTileHeight.Text) / val(lbloriginalheight.Caption)) * 100
110           txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
120           txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")

130       ElseIf optSize(1).value Then
140           txtPercentWidth.Text = GetSetting("PercWidth", CStr(val(txtTileWidth.Text) / val(lbloriginalwidth.Caption)) * 100)
150           txtPercentHeight.Text = GetSetting("PercHeight", CStr(val(txtTileHeight.Text) / val(lbloriginalheight.Caption)) * 100)
160           txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
170           txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")

180           txtTileWidth.Text = Round((val(txtPercentWidth.Text) / 100) * val(lbloriginalwidth.Caption))
190           txtTileHeight.Text = Round((val(txtPercentHeight.Text) / 100) * val(frmResize.lbloriginalheight.Caption))

200       End If

210       chkmaintainratio.value = CInt(GetSetting("MaintainRatio", "1"))

220       cmbResizeType.ListIndex = CInt(GetSetting("ResizeType", "0"))


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        cmdCancel_Click
End Sub

Private Sub cmdOK_Click()
          Dim newB As area

          Dim oldtilew As Integer
          Dim oldtileh As Integer
          Dim tilewidth As Integer
          Dim tileheight As Integer

10        oldtilew = parent.sel.getBoundaries.Right - parent.sel.getBoundaries.Left
20        oldtileh = parent.sel.getBoundaries.Bottom - parent.sel.getBoundaries.Top


30        tilewidth = val(txtTileWidth.Text)
40        tileheight = val(txtTileHeight.Text)

50        If tilewidth = oldtilew And tileheight = oldtileh Then
60            Call cmdCancel_Click
70            Exit Sub
80        End If
          
90        frmGeneral.IsBusy("frmResize.cmdOK_Click") = True
          
100       newB.Left = parent.sel.getBoundaries.Left
110       newB.Top = parent.sel.getBoundaries.Top
120       newB.Right = newB.Left + tilewidth
130       newB.Bottom = newB.Top + tileheight

140       If newB.Right > 1023 Then
150           newB.Left = newB.Left - (newB.Right - 1023)
160           newB.Right = 1023
170       End If

180       If newB.Bottom > 1023 Then
190           newB.Top = newB.Top - (newB.Bottom - 1023)
200           newB.Bottom = 1023
210       End If

220       Call parent.sel.Resize(newB.Left, newB.Top, newB.Right, newB.Bottom, cmbResizeType.ListIndex, True)

230       Call SetSetting("ResizeType", cmbResizeType.ListIndex)

240       If optSize(0).value Then
250           Call SetSetting("OptSize", "0")
260       ElseIf optSize(1).value Then
270           Call SetSetting("OptSize", "1")
280       End If

290       If optSize(0).value Then
300           Call SetSetting("TileWidth", txtTileWidth.Text)
310           Call SetSetting("TileHeight", txtTileHeight.Text)
320       ElseIf optSize(1).value Then
330           Call SetSetting("PercWidth", txtPercentWidth.Text)
340           Call SetSetting("PercHeight", txtPercentHeight.Text)
350       End If

360       Call SetSetting("MaintainRatio", chkmaintainratio.value)

370       Call SaveSettings

380       frmGeneral.IsBusy("frmResize.cmdOK_Click") = False
          
390       DoEvents
400       Unload Me

End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
End Sub

Private Sub txtPercentHeight_Change()
10        Call removeDisallowedCharacters(txtPercentHeight, 1, (1024 / val(lbloriginalheight.Caption)) * 100, True)

20        If (val(txtPercentHeight.Text) / 100) * val(lbloriginalheight.Caption) > 1024 Then
30            txtPercentHeight.Text = (1024 / val(lbloriginalheight.Caption)) * 100
40            txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")
50        End If

60        If percHfocus Then
70            txtTileHeight.Text = Round((val(txtPercentHeight.Text) / 100) * val(frmResize.lbloriginalheight.Caption))

80            optSize(1).value = True

90            If chkmaintainratio.value = 1 Then
100               txtPercentWidth.Text = txtPercentHeight.Text
110               txtTileWidth.Text = Round((val(lbloriginalwidth.Caption) * val(txtPercentWidth.Text)) / 100)
120           End If
130       End If


End Sub

Private Sub txtPercentWidth_Change()
10        Call removeDisallowedCharacters(txtPercentWidth, 1, (1024 / val(lbloriginalwidth.Caption)) * 100, True)

20        If Int(val(txtPercentWidth.Text) / 100) * val(lbloriginalwidth.Caption) > 1024 Then
30            txtPercentWidth.Text = Int((1024 / val(lbloriginalwidth.Caption)) * 100)
40            txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
50        End If

60        If percWfocus Then
70            txtTileWidth.Text = Round((val(txtPercentWidth.Text) / 100) * val(lbloriginalwidth.Caption))

80            optSize(1).value = True

90            If chkmaintainratio.value = 1 Then
100               txtPercentHeight.Text = txtPercentWidth.Text
110               txtTileHeight.Text = Round((val(lbloriginalheight.Caption) * val(txtPercentHeight.Text)) / 100)
120           End If
130       End If



End Sub

Private Sub txtTileHeight_Change()
10        Call removeDisallowedCharacters(txtTileHeight, 0, 1024, True)

20        If tileHfocus Then
30            txtPercentHeight.Text = (val(txtTileHeight.Text) / val(lbloriginalheight.Caption)) * 100
40            txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")
50            optSize(0).value = True

60            If chkmaintainratio.value = 1 Then
70                txtPercentWidth.Text = txtPercentHeight.Text
80                txtTileWidth.Text = Round((val(lbloriginalwidth.Caption) * val(txtPercentWidth.Text)) / 100)
90            End If

100       End If
End Sub

Private Sub txtTileWidth_Change()
10        Call removeDisallowedCharacters(txtTileWidth, 0, 1024, True)

20        If tileWfocus Then
30            txtPercentWidth.Text = (val(txtTileWidth.Text) / val(lbloriginalwidth.Caption)) * 100
40            txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
50            optSize(0).value = True

60            If chkmaintainratio.value = 1 Then
70                txtPercentHeight.Text = txtPercentWidth.Text
80                txtTileHeight.Text = Round((val(lbloriginalheight.Caption) * val(txtPercentHeight.Text)) / 100)
90            End If

100       End If
End Sub

Private Sub txtPercentHeight_GotFocus()
10        tileWfocus = False
20        tileHfocus = False
30        percWfocus = False
40        percHfocus = True
End Sub

Private Sub txtPercentWidth_GotFocus()
10        tileWfocus = False
20        tileHfocus = False
30        percWfocus = True
40        percHfocus = False
End Sub

Private Sub txtTileHeight_GotFocus()
10        tileWfocus = False
20        tileHfocus = True
30        percWfocus = False
40        percHfocus = False
End Sub

Private Sub txtTileWidth_GotFocus()
10        tileWfocus = True
20        tileHfocus = False
30        percWfocus = False
40        percHfocus = False
End Sub

