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
    Set parent = map
End Sub

Private Sub Form_Load()

    Set Me.Icon = frmGeneral.Icon

    Dim oSize As Integer
    oSize = CInt(GetSetting("OptSize", "0"))
    optSize(oSize).value = True

    lbloriginalwidth.Caption = parent.sel.getBoundaries.Right - parent.sel.getBoundaries.Left
    lbloriginalheight.Caption = parent.sel.getBoundaries.Bottom - parent.sel.getBoundaries.Top

    If optSize(0).value Then
        txtTileWidth.Text = GetSetting("TileWidth", lbloriginalwidth.Caption)
        txtTileHeight.Text = GetSetting("TileHeight", lbloriginalheight.Caption)

        txtPercentWidth.Text = (val(txtTileWidth.Text) / val(lbloriginalwidth.Caption)) * 100
        txtPercentHeight.Text = (val(txtTileHeight.Text) / val(lbloriginalheight.Caption)) * 100
        txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
        txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")

    ElseIf optSize(1).value Then
        txtPercentWidth.Text = GetSetting("PercWidth", CStr(val(txtTileWidth.Text) / val(lbloriginalwidth.Caption)) * 100)
        txtPercentHeight.Text = GetSetting("PercHeight", CStr(val(txtTileHeight.Text) / val(lbloriginalheight.Caption)) * 100)
        txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
        txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")

        txtTileWidth.Text = Round((val(txtPercentWidth.Text) / 100) * val(lbloriginalwidth.Caption))
        txtTileHeight.Text = Round((val(txtPercentHeight.Text) / 100) * val(frmResize.lbloriginalheight.Caption))

    End If

    chkmaintainratio.value = CInt(GetSetting("MaintainRatio", "1"))

    cmbResizeType.ListIndex = CInt(GetSetting("ResizeType", "0"))


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click
End Sub

Private Sub cmdOK_Click()
    Dim newB As area

    Dim oldtilew As Integer
    Dim oldtileh As Integer
    Dim tileWidth As Integer
    Dim tileHeight As Integer

    oldtilew = parent.sel.getBoundaries.Right - parent.sel.getBoundaries.Left
    oldtileh = parent.sel.getBoundaries.Bottom - parent.sel.getBoundaries.Top


    tileWidth = val(txtTileWidth.Text)
    tileHeight = val(txtTileHeight.Text)

    If tileWidth = oldtilew And tileHeight = oldtileh Then
        Call cmdCancel_Click
        Exit Sub
    End If
    
    frmGeneral.IsBusy("frmResize.cmdOK_Click") = True
    
    newB.Left = parent.sel.getBoundaries.Left
    newB.Top = parent.sel.getBoundaries.Top
    newB.Right = newB.Left + tileWidth
    newB.Bottom = newB.Top + tileHeight

    If newB.Right > 1023 Then
        newB.Left = newB.Left - (newB.Right - 1023)
        newB.Right = 1023
    End If

    If newB.Bottom > 1023 Then
        newB.Top = newB.Top - (newB.Bottom - 1023)
        newB.Bottom = 1023
    End If

    Call parent.sel.Resize(newB.Left, newB.Top, newB.Right, newB.Bottom, cmbResizeType.ListIndex, True)

    Call SetSetting("ResizeType", cmbResizeType.ListIndex)

    If optSize(0).value Then
        Call SetSetting("OptSize", "0")
    ElseIf optSize(1).value Then
        Call SetSetting("OptSize", "1")
    End If

    If optSize(0).value Then
        Call SetSetting("TileWidth", txtTileWidth.Text)
        Call SetSetting("TileHeight", txtTileHeight.Text)
    ElseIf optSize(1).value Then
        Call SetSetting("PercWidth", txtPercentWidth.Text)
        Call SetSetting("PercHeight", txtPercentHeight.Text)
    End If

    Call SetSetting("MaintainRatio", chkmaintainratio.value)

    Call SaveSettings

    frmGeneral.IsBusy("frmResize.cmdOK_Click") = False
    
    DoEvents
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
End Sub

Private Sub txtPercentHeight_Change()
    Call removeDisallowedCharacters(txtPercentHeight, 0.1, (1024 / val(lbloriginalheight.Caption)) * 100, True)

    If (val(txtPercentHeight.Text) / 100) * val(lbloriginalheight.Caption) > 1024 Then
        txtPercentHeight.Text = (1024 / val(lbloriginalheight.Caption)) * 100
        txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")
    End If

    If percHfocus Then
        txtTileHeight.Text = Round((val(txtPercentHeight.Text) / 100) * val(frmResize.lbloriginalheight.Caption))

        optSize(1).value = True

        If chkmaintainratio.value = 1 Then
            txtPercentWidth.Text = txtPercentHeight.Text
            txtTileWidth.Text = Round((val(lbloriginalwidth.Caption) * val(txtPercentWidth.Text)) / 100)
        End If
    End If


End Sub

Private Sub txtPercentWidth_Change()
    Call removeDisallowedCharacters(txtPercentWidth, 0.1, (1024 / val(lbloriginalwidth.Caption)) * 100, True)

    If Int(val(txtPercentWidth.Text) / 100) * val(lbloriginalwidth.Caption) > 1024 Then
        txtPercentWidth.Text = Int((1024 / val(lbloriginalwidth.Caption)) * 100)
        txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
    End If

    If percWfocus Then
        txtTileWidth.Text = Round((val(txtPercentWidth.Text) / 100) * val(lbloriginalwidth.Caption))

        optSize(1).value = True

        If chkmaintainratio.value = 1 Then
            txtPercentHeight.Text = txtPercentWidth.Text
            txtTileHeight.Text = Round((val(lbloriginalheight.Caption) * val(txtPercentHeight.Text)) / 100)
        End If
    End If



End Sub

Private Sub txtTileHeight_Change()
    Call removeDisallowedCharacters(txtTileHeight, 1, 1024, True)

    If tileHfocus Then
        txtPercentHeight.Text = (val(txtTileHeight.Text) / val(lbloriginalheight.Caption)) * 100
        txtPercentHeight.Text = Format(txtPercentHeight.Text, "Fixed")
        optSize(0).value = True

        If chkmaintainratio.value = 1 Then
            txtPercentWidth.Text = txtPercentHeight.Text
            txtTileWidth.Text = Round((val(lbloriginalwidth.Caption) * val(txtPercentWidth.Text)) / 100)
        End If

    End If
End Sub

Private Sub txtTileWidth_Change()
    Call removeDisallowedCharacters(txtTileWidth, 1, 1024, True)

    If tileWfocus Then
        txtPercentWidth.Text = (val(txtTileWidth.Text) / val(lbloriginalwidth.Caption)) * 100
        txtPercentWidth.Text = Format(txtPercentWidth.Text, "Fixed")
        optSize(0).value = True

        If chkmaintainratio.value = 1 Then
            txtPercentHeight.Text = txtPercentWidth.Text
            txtTileHeight.Text = Round((val(lbloriginalheight.Caption) * val(txtPercentHeight.Text)) / 100)
        End If

    End If
End Sub

Private Sub txtPercentHeight_GotFocus()
    tileWfocus = False
    tileHfocus = False
    percWfocus = False
    percHfocus = True
End Sub

Private Sub txtPercentWidth_GotFocus()
    tileWfocus = False
    tileHfocus = False
    percWfocus = True
    percHfocus = False
End Sub

Private Sub txtTileHeight_GotFocus()
    tileWfocus = False
    tileHfocus = True
    percWfocus = False
    percHfocus = False
End Sub

Private Sub txtTileWidth_GotFocus()
    tileWfocus = True
    tileHfocus = False
    percWfocus = False
    percHfocus = False
End Sub

