VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTilesetEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tileset Editor"
   ClientHeight    =   7650
   ClientLeft      =   540
   ClientTop       =   990
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picdragdrop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   8250
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   3750
      Visible         =   0   'False
      Width           =   240
      Begin VB.Shape dragsel 
         BorderColor     =   &H0000FFFF&
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picgenWallTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   6840
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      Top             =   2040
      Width           =   960
   End
   Begin VB.CommandButton cmdGenWallTileFromSource 
      Caption         =   "Generate Walltiles"
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Generates a walltile set from a vertical tile"
      Top             =   1680
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar ToolbarSource 
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   3240
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Load Source Image"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Drop"
            Object.ToolTipText     =   "Drop Source Image"
            ImageKey        =   "Cancel"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Grid"
            Object.ToolTipText     =   "Show Grid"
            ImageKey        =   "Grid"
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "SelectAll"
            Object.ToolTipText     =   "Select All"
            ImageKey        =   "SelectAll"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "SelectNone"
            Object.ToolTipText     =   "Select None"
            ImageKey        =   "SelectNone"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   7395
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Editor"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox pictemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   8640
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtSnap 
      Height          =   285
      Left            =   6840
      MaxLength       =   3
      TabIndex        =   25
      Text            =   "16"
      Top             =   6720
      Width           =   495
   End
   Begin VB.HScrollBar Hscr 
      Height          =   255
      LargeChange     =   16
      Left            =   240
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.VScrollBar Vscr 
      Height          =   2535
      LargeChange     =   16
      Left            =   4800
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pictarget 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   5280
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   24
      Top             =   5640
      Width           =   960
      Begin VB.Line lineTarget2 
         BorderColor     =   &H000080FF&
         Visible         =   0   'False
         X1              =   32
         X2              =   32
         Y1              =   26
         Y2              =   38
      End
      Begin VB.Line lineTarget1 
         BorderColor     =   &H000080FF&
         Visible         =   0   'False
         X1              =   26
         X2              =   38
         Y1              =   32
         Y2              =   32
      End
   End
   Begin VB.CommandButton cmdRevertAll 
      Caption         =   "Revert Tileset"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdRevertTile 
      Caption         =   "Revert Tile"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox pictilepreviewleft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   5040
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   600
      Width           =   960
   End
   Begin VB.PictureBox picsrcpreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   5280
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   19
      Top             =   4560
      Width           =   960
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   960
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdEditLeft 
      Caption         =   "Edit Tile"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox pictileset_original 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   8520
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox picsource_original 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   8880
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox pictileset 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   240
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   2
      Top             =   600
      Width           =   4560
      Begin VB.Shape leftsel 
         BorderColor     =   &H000000FF&
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picsource 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2400
      Left            =   240
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   22
      Top             =   4440
      Width           =   4560
      Begin VB.Shape shpcursor 
         BorderColor     =   &H000080FF&
         Height          =   240
         Left            =   1080
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Line lineH 
         BorderColor     =   &H000080FF&
         Visible         =   0   'False
         X1              =   104
         X2              =   120
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Line lineV 
         BorderColor     =   &H000080FF&
         Visible         =   0   'False
         X1              =   104
         X2              =   104
         Y1              =   88
         Y2              =   104
      End
      Begin VB.Shape srcsel 
         BorderColor     =   &H00FF0000&
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7320
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.Toolbar ToolbarTop 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Tileset"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Tileset"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Grid"
            Object.ToolTipText     =   "Show Grid"
            ImageKey        =   "Grid"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mirror"
            Object.ToolTipText     =   "Flip Selection Horizontally"
            ImageKey        =   "Mirror"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Flip"
            Object.ToolTipText     =   "Flip Selection Vertically"
            ImageKey        =   "Flip"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ilImages 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":0000
            Key             =   "Ellipse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":0114
            Key             =   "NotRMode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":0468
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":0580
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":0698
            Key             =   "Region"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":09EC
            Key             =   "Filled"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":0D40
            Key             =   "Flip"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1094
            Key             =   "Mirror"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":13E8
            Key             =   "NoClip"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":173C
            Key             =   "Clip"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1A90
            Key             =   "R180"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1DE4
            Key             =   "RLeft"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2138
            Key             =   "RRight"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":248C
            Key             =   "UnFilled"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":27E0
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":28F4
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2A08
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2B1C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2C30
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2D44
            Key             =   "New"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2E58
            Key             =   "Rectangle"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":2F6C
            Key             =   "Arc"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":3080
            Key             =   "RMode"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":33D4
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":3728
            Key             =   "Segment"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":3A7C
            Key             =   "Zoom"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":3DD0
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":4124
            Key             =   "TextDef"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7200
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   65280
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":4478
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":4A12
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":4FAC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":5346
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":58E0
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":5C7A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":6014
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":65AE
            Key             =   "Paste Under"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":6948
            Key             =   "Paste Normal"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":6CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":727C
            Key             =   "Paste Transparent"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":7616
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":7BB0
            Key             =   "Mirror"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":814A
            Key             =   "Flip"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":86E4
            Key             =   "Rotate"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":8C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":9218
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":97B2
            Key             =   "Selection"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":9D4C
            Key             =   "Magnifier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":A0E6
            Key             =   "Dropper"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":A680
            Key             =   "Brush"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":AC1A
            Key             =   "Bucket"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":AFB4
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":B54E
            Key             =   "Rectangle"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":BAE8
            Key             =   "FilledRectangle"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":C082
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":C61C
            Key             =   "Ellipse"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":CBB6
            Key             =   "FilledEllipse"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":D150
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":D6EA
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":DC84
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":E01E
            Key             =   "Eraser"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":E3B8
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":E952
            Key             =   "Grid"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":ECEC
            Key             =   "Tilenr"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":F086
            Key             =   "Hand"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":F620
            Key             =   "SpLine"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":F9BA
            Key             =   "FillInScreen"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":FD54
            Key             =   "TTM"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":100EE
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":10488
            Key             =   "PTM"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":10822
            Key             =   "Airbrush"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":10BBC
            Key             =   "TilenrSelected"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":10F56
            Key             =   "GridSelected"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":112F0
            Key             =   "FillInScreenSelected"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1168A
            Key             =   "FillInShapesSelected"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":11A24
            Key             =   "Paste UnderSelected"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":11DBE
            Key             =   "Paste NormalSelected"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":12158
            Key             =   "Tip"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":124F2
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1288C
            Key             =   "Count"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":12C26
            Key             =   "TileHorizontally"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":12FC0
            Key             =   "TileVertically"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1335A
            Key             =   "Cascade"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":136F4
            Key             =   "MagicWand"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":13A8E
            Key             =   "WallTiles"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":13E28
            Key             =   "WallTilesSelected"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":141C2
            Key             =   "Paste TransparentSelected"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1455C
            Key             =   "ConvertToWalltiles"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":148F6
            Key             =   "CheckForUpdates"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":14C90
            Key             =   "ImportTileset"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1502A
            Key             =   "ExportTileset"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":153C4
            Key             =   "DiscardTileset"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1575E
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":15AF8
            Key             =   "SelectAdd"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":15E92
            Key             =   "SelectRemove"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1622C
            Key             =   "Resize"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":165C6
            Key             =   "CenterInScreen"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":16960
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":16CFA
            Key             =   "SelectNone"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":17094
            Key             =   "SelectLeftTile"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":1742E
            Key             =   "SelectRightTile"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":177C8
            Key             =   "SelectAddLeftTile"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":17B62
            Key             =   "SelectAddRightTile"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":17EFC
            Key             =   "SelectRemoveLeftTile"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":18296
            Key             =   "SelectRemoveRightTile"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":18630
            Key             =   "EditTileset"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTilesetEditor.frx":189CA
            Key             =   "RotateLeft"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblBackground 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5160
      MousePointer    =   2  'Cross
      TabIndex        =   18
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Background color:"
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblSourcePath 
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   3870
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "(Enter ""16"" to Snap on Tiles)"
      Height          =   255
      Left            =   5160
      TabIndex        =   30
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   560
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Shape shptarget 
      BorderColor     =   &H000080FF&
      Height          =   1200
      Left            =   5160
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Label lblSnap 
      Caption         =   "Snap Cursor to (pixels):"
      Height          =   255
      Left            =   5160
      TabIndex        =   26
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label lblCoord2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Snap: 0, 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   28
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lblCoord1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mouse Position: 0, 0"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label lblSrc 
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label lblLeft 
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   960
      Width           =   960
   End
   Begin VB.Shape shpcur 
      BorderColor     =   &H00FF0000&
      Height          =   1200
      Left            =   5160
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Shape shpleft 
      BorderColor     =   &H000000FF&
      Height          =   1095
      Left            =   4920
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblSource 
      Caption         =   "Source Image"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3630
      Width           =   4695
   End
   Begin VB.Label lblTileset 
      Caption         =   "Tileset"
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.Menu mnuTileset 
      Caption         =   "&Tileset"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Tileset"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Tileset &As..."
      End
      Begin VB.Menu mnuSaveAndUse 
         Caption         =   "Use Tileset"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Tileset..."
      End
      Begin VB.Menu mnuOpenRecent 
         Caption         =   "Open Recent Tileset"
         Begin VB.Menu mnusep2 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSource 
      Caption         =   "&Source"
      Begin VB.Menu mnuLoadSource 
         Caption         =   "&Load Image..."
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "&Clear Source"
      End
      Begin VB.Menu mnuFitToTiles 
         Caption         =   "&Resize Image To Fit Tiles"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditTile 
         Caption         =   "&Edit Tile"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRevertTile 
         Caption         =   "&Revert Tile"
      End
      Begin VB.Menu mnuRevertAll 
         Caption         =   "Revert &Tileset"
      End
   End
End
Attribute VB_Name = "frmTilesetEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tilesetpath As String

Dim background As Long

Dim changed As Boolean

Dim startdragX As Single
Dim startdragY As Single

Dim dragging As Boolean


Dim sourceloaded As Boolean
Dim sourcepath As String

Dim sourceselection As area

Dim ignoremouse As Boolean

Private Sub cmdEditLeft_Click()
    If EditImage(Me, pictileset.hDc, leftsel.Left, leftsel.Top, leftsel.width, leftsel.Height, True) Then
        changed = True
        pictileset.Refresh
    End If
    
'    picTemp.Cls
'    picTemp.width = leftsel.width
'    picTemp.height = leftsel.height
'    BitBlt picTemp.hdc, 0, 0, picTemp.width, picTemp.height, pictileset.hdc, leftsel.Left, leftsel.Top, vbSrcCopy
'    picTemp.refresh
'    frmTileEditor.Show vbModal, frmTilesetEditor

    Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub







Private Sub cmdGenWallTileFromSource_Click()
    GenerateWallTile
End Sub
 
Private Sub cmdRevertAll_Click()
    If MessageBox("Are you sure you want to revert this tileset to its original state? All changes to it will be lost.", vbYesNo + vbQuestion, "Revert Tileset") = vbYes Then
        pictileset.Cls
        BitBlt pictileset.hDc, 0, 0, pictileset.width, pictileset.Height, pictileset_original.hDc, 0, 0, vbSrcCopy
        pictileset.Refresh
    End If

    changed = False

End Sub

Private Sub cmdRevertTile_Click()
    BitBlt pictileset.hDc, leftsel.Left, leftsel.Top, leftsel.width, leftsel.Height, pictileset_original.hDc, leftsel.Left, leftsel.Top, vbSrcCopy
    pictileset.Refresh
End Sub













Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    ToolbarSource.width = Me.width
    Line1.x1 = 0
    Line1.x2 = Me.width

    'place color rectangles around the preview pictureboxes
    PlaceRectangleAround shpleft, pictilepreviewleft
    PlaceRectangleAround shpcur, picsrcpreview
    PlaceRectangleAround shptarget, pictarget

    'place and size the scrollbars correctly
    Vscr.Left = picsource.Left + picsource.width
    Vscr.Top = picsource.Top
    Vscr.Height = picsource.Height
    Hscr.Left = picsource.Left
    Hscr.Top = picsource.Top + picsource.Height
    Hscr.width = picsource.width


    sourceloaded = False
    changed = False

    'check settings
    background = CLng(GetSetting("TilesetEditorBack", RGB(0, 0, 0)))
    UpdateBackground

    txtSnap.Text = GetSetting("TilesetEditorSnap", "16")

    Dim ret As String
    ret = GetSetting("TilesetEditorSource", "")
    If FileExists(ret) Then
        Call LoadSource(ret)
    End If

    BitBlt pictileset_original.hDc, 0, 0, pictileset.width, pictileset.Height, frmGeneral.cTileset.Pic_Tileset.hDc, 0, 0, vbSrcCopy
    pictileset_original.Refresh
    BitBlt pictileset.hDc, 0, 0, pictileset.width, pictileset.Height, pictileset_original.hDc, 0, 0, vbSrcCopy
    pictileset.Refresh

    Call DrawSourcePreview

End Sub








Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetStatus ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If changed = True Then
        If MessageBox("Do you wish to save your tileset and apply it to your active map before closing the editor?", vbYesNo + vbQuestion, "Save changes") = vbYes Then
            If Overwrite8bits Then
                Dim ret As Boolean
                ret = SaveTileset(App.path & "\tmptileset.bmp")
                If ret Then
                    Call ApplyTilesetToMap(App.path & "\tmptileset.bmp")
                    DeleteFile App.path & "\tmptileset.bmp"
                    'Call Kill(App.path & "\tmptileset.bmp")
                    
                Else
                    Cancel = True
                    Exit Sub
                End If
            Else
                Cancel = True
                Exit Sub
            End If
        End If
    End If

    Call SetSetting("TilesetEditorSnap", txtSnap.Text)
    Call SetSetting("TilesetEditorSource", sourcepath)
    Call SetSetting("TilesetEditorBack", CStr(lblBackground.BackColor))
    Call SaveSettings

    Unload Me
End Sub




Private Sub Form_Resize()
    ToolbarSource.width = Me.width
    Line1.x1 = 0
    Line1.x2 = Me.width
End Sub

Private Sub frmTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetStatus "Generate walltiles from a vertical tile and drag it into your tileset."
End Sub



Private Sub lblBackground_Click()
    background = GetColor(Me, lblBackground.BackColor, True, True)
    UpdateBackground

    If sourceloaded Then LoadSource (sourcepath)

    Call DrawSourcePreview
    Call RedrawPreviews
End Sub

Private Sub UpdateBackground()
    lblBackground.BackColor = background
    picsource_original.BackColor = background
    picsource.BackColor = background
    pictilepreviewleft.BackColor = background
    picsrcpreview.BackColor = background
    pictarget.BackColor = background
End Sub

Private Sub RedrawPreviews()

    Call DrawTilePreview(picsource, srcsel, picsrcpreview)
    Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)
    Call PlaceCross(0, 0)
End Sub



Private Sub mnuDrop_Click()
    picsource_original.Picture = Nothing

    picsource_original.Cls

    picsource_original.width = 304
    picsource_original.Height = 160
    picsource_original.Refresh
    picsource.Cls
    sourceloaded = False
    sourcepath = ""
    Call DrawSourcePreview

    lblSrc.Caption = ""

    picsrcpreview.Cls
    picsrcpreview.Refresh
    pictarget.Cls
    pictarget.Refresh
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFitToTiles_Click()
    picsource_original.width = picsource_original.width + (TILEW - (picsource_original.width Mod TILEW)) Mod TILEW
    picsource_original.Height = picsource_original.Height + (TILEW - (picsource_original.Height Mod TILEW)) Mod TILEW

    Call DrawSourcePreview
End Sub

Private Sub mnuLoadSource_Click()
    On Error GoTo errh

    ignoremouse = True

    'opens a common dialog
    cd.DialogTitle = "Select an image to import"
    cd.flags = cdlOFNHideReadOnly

    cd.InitDir = GetLastDialogPath("TilesetSourceImage")
    
    cd.Filter = "Supported image files (*.lvl, *.bmp, *.png, *.gif, *.jpg)|*.lvl; *.bmp; *.bm2; *.gif; *.jpg; *.jpeg; *.png|Tilesets (*.lvl, *.bmp)|*.lvl; *.bmp|All files (*.*)|*.*"
    cd.ShowOpen
    
    If cd.filename <> "" Then
        Call LoadSource(cd.filename)

        Call SetLastDialogPath("TilesetSourceImage", GetPathTo(cd.filename))
    End If
    
    ignoremouse = False

    On Error GoTo 0
    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    End If
    HandleError Err, "frmTilesetEditor.mnuLoadSource_Click"
End Sub

Private Sub mnuOpen_Click()
    If changed = True Then
        If MessageBox("Opening another tileset will cause all your changes to the current tileset to be lost, do you wish to save your tileset first?", vbYesNo + vbQuestion, "Save changes") = vbYes Then
            Call SaveTileset(tilesetpath)
        End If
    End If

    'import a tileset
    Call ImportTileset

End Sub

Private Sub mnuSave_Click()
    Call SaveTileset(tilesetpath)
End Sub

Private Sub mnuSaveAndUse_Click()
    Dim ret As Boolean
    ret = SaveTileset(App.path & "\tmptileset.bmp")
    If ret Then
        Call ApplyTilesetToMap(App.path & "\tmptileset.bmp")
        DeleteFile App.path & "\tmptileset.bmp"
        'Call Kill(App.path & "\tmptileset.bmp")
    Else
        MessageBox "Save Failed!"
        Exit Sub
    End If

    Unload Me
End Sub

Private Sub mnuSaveAs_Click()
    Call SaveTileset("")
End Sub

Private Sub picdragdrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragging = True Then
        SetStatus "Drop these tiles on your tileset to swap them. Hold the Control key to copy the tiles."
    Else
        SetStatus "Drop these tiles on your tileset to use them."
    End If
End Sub





Private Sub picgenWallTile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    If Button Then
        picdragdrop.width = picgenWallTile.width
        picdragdrop.Height = picgenWallTile.Height
        dragsel.width = picdragdrop.width
        dragsel.Height = picdragdrop.Height
        picdragdrop.Left = X + picgenWallTile.Left - picgenWallTile.width \ 2
        picdragdrop.Top = Y + picgenWallTile.Top - picgenWallTile.Height \ 2
        picdragdrop.Cls
        BitBlt picdragdrop.hDc, 0, 0, picdragdrop.width, picdragdrop.Height, picgenWallTile.hDc, 0, 0, vbSrcCopy
        picdragdrop.Refresh
        picdragdrop.visible = True
    End If
End Sub

Private Sub picgenWallTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub
    
    SetStatus "Generate walltiles from a vertical tile, then drag them on your tileset to use them."
    If Button Then
        Dim dropX As Single
        Dim dropY As Single
    
        dropX = X + picgenWallTile.Left
        dropY = Y + picgenWallTile.Top
    
        If dropX >= pictileset.Left And dropX <= pictileset.Left + pictileset.width _
           And dropY >= pictileset.Top And dropY <= pictileset.Top + pictileset.Height Then
            'We are dragging the source over the tileset
            'We must snap it on tiles
            picdragdrop.Left = ((dropX - (picdragdrop.width \ 2) - pictileset.Left) \ TILEW) * TILEW + pictileset.Left
            picdragdrop.Top = ((dropY - (picdragdrop.Height \ 2) - pictileset.Top) \ TILEW) * TILEW + pictileset.Top
        Else
            picdragdrop.Left = dropX - (picdragdrop.width \ 2)
            picdragdrop.Top = dropY - (picdragdrop.Height \ 2)
        End If
    End If
End Sub

Private Sub picgenWallTile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    If Button Then
        If X + picgenWallTile.Left >= pictileset.Left And X + picgenWallTile.Left <= pictileset.Left + pictileset.width _
           And Y + picgenWallTile.Top >= pictileset.Top And Y + picgenWallTile.Top <= pictileset.Top + pictileset.Height Then
            'Tiles can be dropped on tileset
            BitBlt pictileset.hDc, picdragdrop.Left - pictileset.Left, picdragdrop.Top - pictileset.Top, picdragdrop.width, picdragdrop.Height, picdragdrop.hDc, 0, 0, vbSrcCopy
            pictileset.Refresh
            changed = True
        End If
        picdragdrop.visible = False
        picdragdrop.Left = -1000
        picdragdrop.Top = -1000
    End If
End Sub

Private Sub picsource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    If Button = vbRightButton Then
        If X + picsource.Left >= pictileset.Left And X + picsource.Left <= pictileset.Left + pictileset.width _
           And Y + picsource.Top >= pictileset.Top And Y + picsource.Top <= pictileset.Top + pictileset.Height Then
            'Tiles can be dropped on tileset
            BitBlt pictileset.hDc, picdragdrop.Left - pictileset.Left, picdragdrop.Top - pictileset.Top, picdragdrop.width, picdragdrop.Height, picdragdrop.hDc, 0, 0, vbSrcCopy
            pictileset.Refresh
            changed = True
        End If
        picdragdrop.visible = False
        picdragdrop.Left = -1000
        picdragdrop.Top = -1000
    End If
End Sub




Private Sub ToolbarSource_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Open"
        Call mnuLoadSource_Click
    Case "Drop"
        Call mnuDrop_Click
    End Select
End Sub

Private Sub ToolbarTop_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Open"
        Call mnuOpen_Click
    Case "Save"
        Call mnuSave_Click
        
    Case "Mirror"
        Call FlipH(pictileset, leftsel.Left, leftsel.width, leftsel.Top, leftsel.Height)
    
    Case "Flip"
        Call FlipV(pictileset, leftsel.Left, leftsel.width, leftsel.Top, leftsel.Height)
    End Select
End Sub

Private Sub txtSnap_Change()
    Call removeDisallowedCharacters(txtSnap, 1, picsource_original.width \ 2, False)
End Sub


Private Sub Vscr_Scroll()
    Call DrawSourcePreview

End Sub


Private Sub Hscr_Scroll()
    Call DrawSourcePreview
End Sub

Private Sub picsource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    If Button = vbLeftButton Then
        startdragX = X + Hscr.value
        startdragY = Y + Vscr.value

        Call picsource_MouseMove(Button, Shift, X, Y)
    ElseIf Button = vbRightButton Then
        picdragdrop.width = srcsel.width
        picdragdrop.Height = srcsel.Height
        dragsel.width = picdragdrop.width
        dragsel.Height = picdragdrop.Height
        picdragdrop.Left = X + picsource.Left - picsource.width \ 2
        picdragdrop.Top = Y + picsource.Top - picsource.Height \ 2
        picdragdrop.Cls
        BitBlt picdragdrop.hDc, 0, 0, picdragdrop.width, picdragdrop.Height, picsource_original.hDc, srcsel.Left + Hscr.value, srcsel.Top + Vscr.value, vbSrcCopy
        picdragdrop.Refresh
        picdragdrop.visible = True
    End If
End Sub

Private Sub picsource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub


    If Not sourceloaded Then
        SetStatus "Load an image to import tiles to your tileset."
        Exit Sub
    ElseIf dragging = True Then
        SetStatus "You cannot drag tiles from the tileset to the source image."
    ElseIf picdragdrop.visible = False Then
        SetStatus "Use the left mouse button to select a part of the image, then drag it to your tileset with the right mouse button."
    End If

    If Button = 0 Then
        Call PlaceCross(X, Y)
    Else
        lineV.visible = False
        lineH.visible = False
        shpcursor.visible = False
    End If

    If Button = vbLeftButton Then
        Call MoveSourceShape(X, Y)
    ElseIf Button = vbRightButton Then
        'Drag tile(s) toward tileset
        Dim dropX As Single
        Dim dropY As Single

        dropX = X + picsource.Left
        dropY = Y + picsource.Top

        If dropX >= pictileset.Left And dropX <= pictileset.Left + pictileset.width _
           And dropY >= pictileset.Top And dropY <= pictileset.Top + pictileset.Height Then
            'We are dragging the source over the tileset
            'We must snap it on tiles
            picdragdrop.Left = ((dropX - (picdragdrop.width \ 2) - pictileset.Left) \ TILEW) * TILEW + pictileset.Left
            picdragdrop.Top = ((dropY - (picdragdrop.Height \ 2) - pictileset.Top) \ TILEW) * TILEW + pictileset.Top
        Else
            picdragdrop.Left = dropX - (picdragdrop.width \ 2)
            picdragdrop.Top = dropY - (picdragdrop.Height \ 2)
        End If
        SetStatus "Drop these tiles on your tileset to use them."
    Else
        Exit Sub
    End If
    Call DrawTilePreview(picsource, srcsel, picsrcpreview)
End Sub

Private Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    If ((Shift = 2 Or Shift = 3 Or Shift = 6) Or Button = vbRightButton) Then
        'Start drag and drop
        If X >= leftsel.Left And X <= leftsel.Left + leftsel.width _
           And Y >= leftsel.Top And Y <= leftsel.Top + leftsel.Height Then


            dragging = True
            picdragdrop.width = leftsel.width
            picdragdrop.Height = leftsel.Height
            dragsel.width = picdragdrop.width
            dragsel.Height = picdragdrop.Height
            picdragdrop.Cls
            BitBlt picdragdrop.hDc, 0, 0, picdragdrop.width, picdragdrop.Height, pictileset.hDc, leftsel.Left + Hscr.value, leftsel.Top + Vscr.value, vbSrcCopy
            picdragdrop.Refresh

            pictemp.Cls
            pictemp.width = picdragdrop.width
            pictemp.Height = picdragdrop.Height
            BitBlt pictemp.hDc, 0, 0, pictemp.width, pictemp.Height, picdragdrop.hDc, 0, 0, vbSrcCopy
            pictemp.Refresh

            Call pictileset_MouseMove(Button, Shift, X, Y)
            picdragdrop.visible = True
        End If
    ElseIf Button = vbLeftButton Then
        startdragX = X
        startdragY = Y

        Call pictileset_MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    If dragging Then
        'Start drag and drop on tileset

        If X >= 0 And X <= pictileset.width _
           And Y >= 0 And Y <= pictileset.Height Then
            'We are dragging the source over the tileset
            'We must snap it on tiles
            picdragdrop.visible = True
            Dim tmpX As Integer
            Dim tmpY As Integer
            tmpX = ((X - (picdragdrop.width \ 2)) \ TILEW) * TILEW
            tmpY = ((Y - (picdragdrop.Height \ 2)) \ TILEW) * TILEW
            If tmpX <= 0 Then tmpX = 0
            If tmpX + picdragdrop.width >= pictileset.width Then tmpX = pictileset.width - picdragdrop.width
            If tmpY <= 0 Then tmpY = 0
            If tmpY + picdragdrop.Height >= pictileset.Height Then tmpY = pictileset.Height - picdragdrop.Height


            picdragdrop.Left = tmpX + pictileset.Left
            picdragdrop.Top = tmpY + pictileset.Top

            If Not (Shift = 2 Or Shift = 3 Or Shift = 6) Then
                'preview the swapped tiles
                BitBlt pictileset.hDc, leftsel.Left, leftsel.Top, picdragdrop.width, picdragdrop.Height, pictileset.hDc, picdragdrop.Left - pictileset.Left, picdragdrop.Top - pictileset.Top, vbSrcCopy
            Else
                'restore original tiles
                BitBlt pictileset.hDc, leftsel.Left, leftsel.Top, pictemp.width, pictemp.Height, pictemp.hDc, 0, 0, vbSrcCopy
            End If
            pictileset.Refresh
        Else
            picdragdrop.visible = False
        End If
        SetStatus "Drop these tiles on your tileset to swap them. Hold the Control key to copy the tiles."

    ElseIf Button = vbLeftButton Then
        Call MoveTilesetShape(X, Y)
        Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)
    End If

End Sub

Private Sub pictileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ignoremouse Then Exit Sub

    startdragX = -1
    startdragY = -1

    If dragging = True Then
        If X >= 0 And X <= pictileset.width _
           And Y >= 0 And Y <= pictileset.Height Then
            'Tiles can be dropped on tileset

            'Swap tiles, unless holding ctrl
            If Not (Shift = 2 Or Shift = 3 Or Shift = 6) Then
                BitBlt pictileset.hDc, leftsel.Left, leftsel.Top, picdragdrop.width, picdragdrop.Height, pictileset.hDc, picdragdrop.Left - pictileset.Left, picdragdrop.Top - pictileset.Top, vbSrcCopy
            End If
            BitBlt pictileset.hDc, picdragdrop.Left - pictileset.Left, picdragdrop.Top - pictileset.Top, picdragdrop.width, picdragdrop.Height, picdragdrop.hDc, 0, 0, vbSrcCopy
            pictileset.Refresh

            leftsel.Left = picdragdrop.Left - pictileset.Left
            leftsel.Top = picdragdrop.Top - pictileset.Top
            changed = True

        End If
        picdragdrop.visible = False
        picdragdrop.Left = -1000
        picdragdrop.Top = -1000
    End If

    dragging = False
End Sub

Private Sub MoveTilesetShape(X As Single, Y As Single)
    Dim Top As Integer
    Dim Bottom As Integer
    Dim Left As Integer
    Dim Right As Integer

    If X < 0 Then X = 0
    If X > pictileset.width Then X = pictileset.width
    If Y < 0 Then Y = 0
    If Y > pictileset.Height Then Y = pictileset.Height

    If startdragX = -1 Or startdragY = -1 Then
        Left = (X \ TILEW) * TILEW
        Right = Left + TILEW
        Top = (Y \ TILEW) * TILEW
        Bottom = Top + TILEW
    Else

        If X < startdragX Then
            Right = ((startdragX + TILEW - 1) \ TILEW) * TILEW
            Left = Right - ((Right - X + TILEW) \ TILEW) * TILEW

            If Right - Left < TILEW Then Left = Right - TILEW
        Else
            Left = (startdragX \ TILEW) * TILEW
            Right = ((X - Left + TILEW) \ TILEW) * TILEW + Left

            If Right - Left < TILEW Then Right = Left + TILEW
        End If

        If Y < startdragY Then
            Bottom = ((startdragY + TILEW - 1) \ TILEW) * TILEW
            Top = Bottom - ((Bottom - Y + TILEW) \ TILEW) * TILEW

            If Bottom - Top < TILEW Then Top = Bottom - TILEW
        Else
            Top = (startdragY \ TILEW) * TILEW
            Bottom = ((Y - Top + TILEW) \ TILEW) * TILEW + Top

            If Bottom - Top < TILEW Then Bottom = Top + TILEW
        End If

    End If

    'Make sure cursor doesnt go out of tileset
    Do While Left < 0
        Left = Left + TILEW
    Loop

    Do While Right > pictileset.width
        Right = Right - TILEW
    Loop

    Do While Top < 0
        Top = Top + TILEW
    Loop

    Do While Bottom > pictileset.Height
        Bottom = Bottom - TILEW
    Loop

    'Adjust selection off the border if needed
    If Right - Left = 0 Then
        If Left < TILEW Then
            Left = 0
        Else
            Left = pictileset.width - TILEW
        End If
        Right = Left + TILEW
    End If
    If Bottom - Top = 0 Then
        If Top < TILEW Then
            Top = 0
        Else
            Top = pictileset.Height - TILEW
        End If
        Bottom = Top + TILEW
    End If

    If leftsel.Left <> Left Then leftsel.Left = Left
    If leftsel.Top <> Top Then leftsel.Top = Top
    If leftsel.width <> Right - Left Then leftsel.width = Right - Left
    If leftsel.Height <> Bottom - Top Then leftsel.Height = Bottom - Top



    lblLeft.Caption = leftsel.width \ TILEW & " " & Chr(215) & " " & leftsel.Height \ TILEW

End Sub

Private Sub MoveSourceShape(X As Single, Y As Single)

    If Not sourceloaded Then Exit Sub

    Dim snapwidth As Integer

    Dim realX As Single
    Dim realY As Single

    Dim Top As Integer
    Dim Bottom As Integer
    Dim Left As Integer
    Dim Right As Integer

    snapwidth = val(txtSnap.Text)


    'Scroll if needed
    Call scroll(picsource, Hscr, Vscr, X, Y)

    realX = X + Hscr.value
    realY = Y + Vscr.value

    If startdragX = -1 Or startdragY = -1 Then
        Left = (X \ snapwidth) * snapwidth
        Right = Left + TILEW
        Top = (Y \ snapwidth) * snapwidth
        Bottom = Top + TILEW
    Else

        If realX < startdragX Then
            Right = (startdragX \ snapwidth) * snapwidth
            If snapwidth = 16 Then Right = Right + TILEW
            Left = Right - ((Right - realX + TILEW) \ TILEW) * TILEW

            If Right - Left < TILEW Then Left = Right - TILEW
        Else
            Left = (startdragX \ snapwidth) * snapwidth
            Right = ((realX - Left + TILEW) \ TILEW) * TILEW + Left

            If Right - Left < TILEW Then Right = Left + TILEW
        End If

        If realY < startdragY Then
            Bottom = (startdragY \ snapwidth) * snapwidth
            If snapwidth = 16 Then Bottom = Bottom + TILEW
            Top = Bottom - ((Bottom - realY + TILEW) \ TILEW) * TILEW

            If Bottom - Top < TILEW Then Top = Bottom - TILEW
        Else
            Top = (startdragY \ snapwidth) * snapwidth
            Bottom = ((realY - Top + TILEW) \ TILEW) * TILEW + Top

            If Bottom - Top < TILEW Then Bottom = Top + TILEW
        End If

    End If

    'Make sure cursor doesnt go out of tileset
    Do While Left < 0
        Left = Left + TILEW
    Loop

    Do While Right > picsource_original.width
        Right = Right - TILEW
    Loop

    Do While Top < 0
        Top = Top + TILEW
    Loop

    Do While Bottom > picsource_original.Height
        Bottom = Bottom - TILEW
    Loop

    'Adjust selection off the border if needed
    If Right - Left = 0 Then
        If Left < TILEW Then
            Left = 0
        Else
            Left = picsource_original.width - TILEW
        End If
        Right = Left + TILEW
    End If
    If Bottom - Top = 0 Then
        If Top < TILEW Then
            Top = 0
        Else
            Top = picsource_original.Height - TILEW
        End If
        Bottom = Top + TILEW
    End If

    sourceselection.Left = Left
    sourceselection.Right = Right
    sourceselection.Top = Top
    sourceselection.Bottom = Bottom

    Left = Left - Hscr.value
    Right = Right - Hscr.value
    Top = Top - Vscr.value
    Bottom = Bottom - Vscr.value

    If srcsel.Left <> Left Then srcsel.Left = Left
    If srcsel.Top <> Top Then srcsel.Top = Top
    If srcsel.width <> Right - Left Then srcsel.width = Right - Left
    If srcsel.Height <> Bottom - Top Then srcsel.Height = Bottom - Top

    If lblCoord1.Caption <> "Top-left: " & sourceselection.Left & ", " & sourceselection.Top Then
        lblCoord1.Caption = "Top-left: " & sourceselection.Left & ", " & sourceselection.Top
    End If
    If lblCoord2.Caption <> "Bottom-right: " & sourceselection.Right & ", " & sourceselection.Bottom Then
        lblCoord2.Caption = "Bottom-right: " & sourceselection.Right & ", " & sourceselection.Bottom
    End If

    lblSrc.Caption = srcsel.width \ TILEW & " " & Chr(215) & " " & srcsel.Height \ TILEW
End Sub

'Private Sub Swap(X As Single, Y As Single)
'    Dim tmp As Integer
'    tmp = X
'    X = Y
'    Y = tmp
'
'End Sub

Private Sub DrawTilePreview(ByRef srcPic As PictureBox, ByRef srcshape As Shape, ByRef destPic As PictureBox)

    Call DrawImagePreview(srcPic, srcshape, destPic, Shape1, destPic.BackColor)
    
'    destPic.Cls
'    If srcshape.Width > srcshape.Height Then
'        'Resize considering width
'
'        If srcshape.Width = destPic.Width Then
'            'Same size, use bitblt
'            BitBlt destPic.hDC, 0, (destPic.Height \ 2) - (srcshape.Height \ 2), srcshape.Width, srcshape.Height, srcPic.hDC, srcshape.Left, srcshape.Top, vbSrcCopy
'        ElseIf srcshape.Width < destPic.Width Then
'            'Source is smaller, use pixel resize
'            SetStretchBltMode destPic.hDC, COLORONCOLOR
'            StretchBlt destPic.hDC, 0, (destPic.Height \ 2) - ((srcshape.Height / (srcshape.Width / destPic.Width)) \ 2), destPic.Width, srcshape.Height / (srcshape.Width / destPic.Width), srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.Width, srcshape.Height, vbSrcCopy
'        Else
'            'Source is larger, use halftone resize
'            SetStretchBltMode destPic.hDC, HALFTONE
'            StretchBlt destPic.hDC, 0, (destPic.Height \ 2) - ((srcshape.Height / (srcshape.Width / destPic.Width)) \ 2), destPic.Width, srcshape.Height / (srcshape.Width / destPic.Width), srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.Width, srcshape.Height, vbSrcCopy
'        End If
'    Else
'        If srcshape.Height = destPic.Height Then
'            'Same size, use bitblt
'            BitBlt destPic.hDC, (destPic.Width \ 2) - (srcshape.Width \ 2), 0, srcshape.Width, srcshape.Height, srcPic.hDC, srcshape.Left, srcshape.Top, vbSrcCopy
'        ElseIf srcshape.Height < destPic.Height Then
'            'Source is smaller, use pixel resize
'            SetStretchBltMode destPic.hDC, COLORONCOLOR
'            StretchBlt destPic.hDC, (destPic.Width \ 2) - ((srcshape.Width / (srcshape.Height / destPic.Height)) \ 2), 0, srcshape.Width / (srcshape.Height / destPic.Height), destPic.Height, srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.Width, srcshape.Height, vbSrcCopy
'        Else
'            'Source is larger, use halftone resize
'            SetStretchBltMode destPic.hDC, HALFTONE
'            StretchBlt destPic.hDC, (destPic.Width \ 2) - ((srcshape.Width / (srcshape.Height / destPic.Height)) \ 2), 0, srcshape.Width / (srcshape.Height / destPic.Height), destPic.Height, srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.Width, srcshape.Height, vbSrcCopy
'        End If
'    End If
'    destPic.Refresh

End Sub

Private Sub DrawSourcePreview()
    picsource.Cls

    If picsource_original.width < 304 Then
        picsource.width = picsource_original.width
    Else
        picsource.width = 304
    End If
    If picsource_original.Height < 160 Then
        picsource.Height = picsource_original.Height
    Else
        picsource.Height = 160
    End If

    If picsource_original.width > picsource.width Then
        Hscr.visible = True
        Hscr.Min = 0
        Hscr.Max = picsource_original.width - picsource.width
    Else
        Hscr.value = 0
        Hscr.visible = False
    End If
    If picsource_original.Height > picsource.Height Then
        Vscr.visible = True
        Vscr.Min = 0
        Vscr.Max = picsource_original.Height - picsource.Height
    Else
        Vscr.value = 0
        Vscr.visible = False
    End If

    If sourceloaded = True Then


        picsource.BorderStyle = 0

        lblSource.Caption = "Source Image - " & GetFileTitle(sourcepath) & " (" & picsource_original.width & " " & Chr(215) & " " & picsource_original.Height & ")"
        lblSourcePath.Caption = sourcepath

        If picsource_original.width = 304 And picsource_original.Height = 160 Then
            'consider it is a tileset, set the snap automatically to tile-snap
            txtSnap.Text = "16"
        End If



        If picsource_original.width Mod TILEW <> 0 Or picsource_original.Height Mod TILEW <> 0 Then
            mnuFitToTiles.Enabled = True
        Else
            mnuFitToTiles.Enabled = False
        End If


        srcsel.Left = sourceselection.Left - Hscr.value
        srcsel.Top = sourceselection.Top - Vscr.value
        srcsel.width = sourceselection.Right - sourceselection.Left
        srcsel.Height = sourceselection.Bottom - sourceselection.Top

        srcsel.visible = True


        BitBlt picsource.hDc, 0, 0, picsource.width, picsource.Height, picsource_original.hDc, Hscr.value, Vscr.value, vbSrcCopy
    Else
        picsource.BorderStyle = 1
        lblSource.Caption = "No Source Image Loaded"
        lblSourcePath.Caption = ""

        lblCoord1.Caption = ""
        lblCoord2.Caption = ""
        srcsel.visible = False
        lineH.visible = False
        lineV.visible = False
        Hscr.visible = False
        Vscr.visible = False

        picsource.Cls

        mnuFitToTiles.Enabled = False
    End If
    picsource.Refresh
End Sub

Friend Sub InitLeftSelection(selection As TilesetSelection)
    Call SetLeftSelection(selection.tilenr, CInt(selection.tileSize.X), CInt(selection.tileSize.Y))
End Sub


Private Sub SetLeftSelection(tilenr As Integer, sizeX As Integer, sizeY As Integer)
'Sets the left selection of the tileset on the given tilenr
    If tilenr > 190 Or tilenr + (sizeX - 1) > 190 Or tilenr + ((sizeY - 1) * 19) > 190 Then
        tilenr = 1
        sizeX = 1
        sizeY = 1
    End If

    'move the shape
    leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
    leftsel.Top = ((tilenr - 1) \ 19) * TILEW

    leftsel.width = TILEW * sizeX
    leftsel.Height = TILEW * sizeY


    lblLeft.Caption = leftsel.width \ TILEW & " " & Chr(215) & " " & leftsel.Height \ TILEW

    'Update the tileset preview
    Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)
End Sub


Private Sub scroll(ByRef curpic As PictureBox, ByRef HScroll As HScrollBar, ByRef VScroll As VScrollBar, X As Single, Y As Single)
    If X < 0 Then
        If Hscr.value - 1 > Hscr.Min Then
            Hscr.value = Hscr.value - 1
        Else
            Hscr.value = Hscr.Min
        End If
    End If
    If X > picsource.width Then
        If Hscr.value + 1 < Hscr.Max Then
            Hscr.value = Hscr.value + 1
        Else
            Hscr.value = Hscr.Max
        End If
    End If
    If Y < 0 Then
        If Vscr.value - 1 > Vscr.Min Then
            Vscr.value = Vscr.value - 1
        Else
            Vscr.value = Vscr.Min
        End If
    End If
    If Y > picsource.Height Then
        If Vscr.value + 1 < Vscr.Max Then
            Vscr.value = Vscr.value + 1
        Else
            Vscr.value = Vscr.Max
        End If
    End If
End Sub

Private Sub PlaceCross(refx As Single, refy As Single)
    Dim X As Single
    Dim Y As Single
    Dim snap As Integer

    If Not sourceloaded Then
        shpcursor.visible = False
        srcsel.visible = False
        lineV.visible = False
        lineH.visible = False
        lineTarget1.visible = False
        lineTarget2.visible = False
        Exit Sub
    End If
    snap = val(txtSnap.Text)

    X = ((refx + Hscr.value) \ snap) * snap - Hscr.value
    Y = ((refy + Vscr.value) \ snap) * snap - Vscr.value


    If snap = 16 Then
        'Snapping to tiles, which means selected tile is included in selection,
        'no matter which way selection goes

        If X + Hscr.value + TILEW > picsource_original.width Then X = picsource_original.width - TILEW - Hscr.value
        If Y + Vscr.value + TILEW > picsource_original.Height Then Y = picsource_original.Height - TILEW - Vscr.value

        lineV.visible = False
        lineH.visible = False
        lineTarget1.visible = False
        lineTarget2.visible = False

        shpcursor.visible = True
        shpcursor.Left = X
        shpcursor.Top = Y

        pictarget.Cls
        SetStretchBltMode pictarget.hDc, COLORONCOLOR
        StretchBlt pictarget.hDc, 0, 0, pictarget.width, pictarget.Height, picsource_original.hDc, X + Hscr.value, Y + Vscr.value, TILEW, TILEW, vbSrcCopy
        pictarget.Refresh
    Else
        shpcursor.visible = False
        lineV.visible = True
        lineH.visible = True
        lineTarget1.visible = True
        lineTarget2.visible = True

        lineV.x1 = X
        lineV.x2 = X

        lineV.y1 = Y - 8
        lineV.y2 = Y + 8

        lineH.x1 = X - 8
        lineH.x2 = X + 8
        lineH.y1 = Y
        lineH.y2 = Y
        Dim zoomlevel As Integer
        zoomlevel = 8
        pictarget.Cls
        SetStretchBltMode pictarget.hDc, COLORONCOLOR
        StretchBlt pictarget.hDc, 0, 0, pictarget.width, pictarget.Height, picsource_original.hDc, X + Hscr.value - zoomlevel / 2, Y + Vscr.value - zoomlevel / 2, zoomlevel, zoomlevel, vbSrcCopy
        pictarget.Refresh
    End If

    lblCoord1.Caption = "Mouse Position: " & refx + Hscr.value & ", " & refy + Vscr.value
    lblCoord2.Caption = "Snap: " & X + Hscr.value & ", " & Y + Vscr.value

End Sub

Private Sub PlaceRectangleAround(ByRef refshape As Shape, ByRef refpic As PictureBox)
    refshape.Left = refpic.Left - 1
    refshape.Top = refpic.Top - 1
    refshape.width = refpic.width + 2
    refshape.Height = refpic.Height + 2
End Sub


Private Sub SetStatus(Text As String)
    StatusBar.SimpleText = Text
End Sub

Private Function SaveTileset(p As String)

    On Error GoTo errh
    Dim path As String

    
    If p = "" Or p = "Default" Or GetExtension(p) = "lvl" Then
        'no path is given, show dialog
        cd.Filter = "*.bmp|*.bmp"
        cd.DialogTitle = "Save Tileset"
        'ask for overwrite

        If tilesetpath <> "Default" Then
            cd.InitDir = tilesetpath
            cd.filename = GetFileNameWithoutExtension(tilesetpath) & ".bmp"
        Else
            cd.filename = "*.bmp"
        End If
        cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

        'if current tileset is from a .lvl, replace with .bmp for default name


        cd.ShowSave

        path = cd.filename
    Else
        path = p
    End If

    pictileset.Refresh
    pictileset_original.Cls
    BitBlt pictileset_original.hDc, 0, 0, pictileset.width, pictileset.Height, pictileset.hDc, 0, 0, vbSrcCopy
    pictileset_original.Refresh


    'apply to the picture

    pictileset_original.Picture = pictileset.Image

    'save it
    Call SavePicture(pictileset_original.Picture, path)

    tilesetpath = path

    SaveTileset = True

    lblTileset.Caption = "Tileset - ''" & tilesetpath & "''"
    
    changed = False
    
    Exit Function
errh:
    If Err = cdlCancel Then
        SaveTileset = False
        Exit Function
    End If

    On Error GoTo 0
    Exit Function
End Function


Private Sub ImportTileset()
'Imports a tileset
    On Error GoTo ImportTileset_Error

    On Error GoTo errh

    'opens a common dialog
    cd.DialogTitle = "Select a tileset to import"
    cd.flags = cdlOFNHideReadOnly
    cd.Filter = "Tilesets (*.lvl, *.bmp)|*.lvl; *.bmp|All files (*.*)|*.*"
    cd.ShowOpen

    If GetExtension(cd.filetitle) <> "lvl" And GetExtension(cd.filetitle) <> "bmp" And GetExtension(cd.filetitle) <> "bm2" Then
        If MessageBox("Specified file does not have a .lvl or .bmp extension and might not be a valid tileset, do you wish to load it anyway?", vbYesNo + vbCritical, "Invalid file type") = vbNo Then
            Exit Sub
        End If
    End If

    If GetExtension(cd.filetitle) = "lvl" Then
        If Not frmGeneral.hasLVLaTileset(cd.filename) Then
            'load default tileset
            Call LoadTileset("")
            Exit Sub
        End If
    End If
    'imports the given tileset
    Call LoadTileset(cd.filename)

    changed = True
    
    Exit Sub
errh:
    'if something goes wrong, use the default tileset
    If Err = cdlCancel Then
        Exit Sub
    Else
        MessageBox Err & " " & Err.description, vbCritical
    End If

    On Error GoTo 0
    Exit Sub

ImportTileset_Error:
    HandleError Err, "frmTilesetEditor.ImportTileset"
End Sub

Private Sub LoadTileset(path As String)
    pictileset_original.Cls
    If path = "" Then
        BitBlt pictileset_original.hDc, 0, 0, pictileset.width, pictileset.Height, frmGeneral.picdefaulttileset.hDc, 0, 0, vbSrcCopy
    Else
        pictileset_original.Picture = LoadPicture(path)
    End If
    pictileset_original.Refresh

    pictileset.Cls
    BitBlt pictileset.hDc, 0, 0, pictileset.width, pictileset.Height, pictileset_original.hDc, 0, 0, vbSrcCopy

    pictileset.Refresh

    Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)

    changed = False
End Sub

Private Sub LoadSource(path As String)
    
    Call LoadPic(picsource_original, path)


    sourceloaded = True

    sourcepath = path

    srcsel.Left = 0
    srcsel.Top = 0
    srcsel.width = TILEW
    srcsel.Height = TILEW

    sourceselection.Left = 0
    sourceselection.Right = TILEW
    sourceselection.Top = 0
    sourceselection.Bottom = TILEW

    Call DrawSourcePreview

    Call DrawTilePreview(picsource, srcsel, picsrcpreview)
    pictarget.Cls
    pictarget.Refresh

    Vscr.value = 0
    Hscr.value = 0
End Sub

'Sub UpdateFromTileEditor()
'    BitBlt pictileset.hdc, leftsel.Left, leftsel.Top, leftsel.width, leftsel.height, frmTileEditor.PicCurrent.hdc, 0, 0, vbSrcCopy
'    pictileset.refresh
'
'    changed = True
'End Sub

Private Function Overwrite8bits() As Boolean
    If frmGeneral.ActiveTilesetIs8bit Then
        If MessageBox("Current map tileset has a 8 bits color depth. This operation will replace it by a 24 bits tileset, which is not compatible with all map editors, but fully compatible with Continuum and allows more colors to be displayed. Do you wish to continue?", vbQuestion + vbOKCancel, "Saving tileset...") = vbOK Then
            Overwrite8bits = True
        Else
            Overwrite8bits = False
        End If
    Else
        Overwrite8bits = True
    End If
End Function

Private Sub ApplyTilesetToMap(path As String)
    Call frmGeneral.ApplyEditedTileset(path)
End Sub

Private Sub RotateCW(ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long)
    'Square rotation; width = height
    Dim i As Long, j As Long, halfw As Long
    Dim tmp As Long
    Dim Xmax As Long, Ymax As Long
    
    Xmax = X + nWidth - 1
    Ymax = Y + nWidth - 1
    halfw = nWidth \ 2 - 1
    
    For i = 0 To halfw
        For j = 0 To halfw
            'Store top-left
            tmp = GetPixel(hDc, X + i, Y + j)
            
            'Set top-left
            Call SetPixel(hDc, X + i, Y + j, GetPixel(hDc, X + j, Ymax - i))
            
            'Set bottom-left
            Call SetPixel(hDc, X + j, Ymax - i, GetPixel(hDc, Xmax - i, Ymax - j))
            
            'Set bottom-right
            Call SetPixel(hDc, Xmax - i, Ymax - j, GetPixel(hDc, Xmax - j, Y + i))
            
            'Set top-right
            Call SetPixel(hDc, Xmax - j, Y + i, tmp)
            
        Next
    Next
    
End Sub

Sub GenerateWallTile()
    Dim i As Integer
    Dim j As Integer
    Dim srchDC As Long
    Dim myDC As Long
    
    
    picgenWallTile.Cls
    
    srchDC = pictileset.hDc
    myDC = picgenWallTile.hDc
    
    'Copy vertical tile
    BitBlt myDC, 0, TILEH * 1, TILEW, TILEH, srchDC, leftsel.Left, leftsel.Top, vbSrcCopy
    BitBlt myDC, TILEW * 2, TILEH * 1, TILEW, TILEH, srchDC, leftsel.Left, leftsel.Top, vbSrcCopy
    
    'Rotate 90
    BitBlt myDC, 0, 0, TILEW, TILEH, srchDC, leftsel.Left, leftsel.Top, vbSrcCopy
    Call RotateCW(myDC, 0, 0, TILEW)
    
    'Spread the rotated tile
    BitBlt myDC, TILEW, 0, TILEW, TILEH, myDC, 0, 0, vbSrcCopy
    BitBlt myDC, TILEW * 2, 0, TILEW, TILEH, myDC, 0, 0, vbSrcCopy
    BitBlt myDC, 0, TILEH * 2, TILEW, TILEH, myDC, 0, 0, vbSrcCopy
    BitBlt myDC, TILEW, TILEH * 2, TILEW, TILEH, myDC, 0, 0, vbSrcCopy
    BitBlt myDC, TILEW * 2, TILEH * 2, TILEW, TILEH, myDC, 0, 0, vbSrcCopy
    
    'Corners...
    'Top-left
    For i = 0 To 15
        BitBlt myDC, 0, i, i, 1, myDC, 0, TILEH + i, vbSrcCopy
    Next
    
    'Top-right
    For i = 0 To 15
        BitBlt myDC, TILEW * 2 + i, TILEH - i, 1, i, myDC, TILEW * 2 + i, TILEH * 2 - i, vbSrcCopy
    Next
    
    'Bottom-left
    For i = 0 To 15
        BitBlt myDC, i, TILEH * 2, 1, 16 - i, myDC, i, TILEH, vbSrcCopy
    Next

    'Bottom-right
    For i = 0 To 15
        BitBlt myDC, TILEW * 2 + i, TILEH * 2, 1, i, myDC, TILEW * 2 + i, TILEH, vbSrcCopy
    Next
    
    'Center...
    BitBlt myDC, TILEW, TILEH, 8, 8, myDC, TILEW * 2, TILEH * 2, vbSrcCopy
    BitBlt myDC, TILEW + 8, TILEH + 8, 8, 8, myDC, 8, 8, vbSrcCopy
    BitBlt myDC, TILEW + 8, TILEH, 8, 8, myDC, 8, TILEH * 2, vbSrcCopy
    BitBlt myDC, TILEW, TILEH + 8, 8, 8, myDC, TILEW * 2, 8, vbSrcCopy
    
    
    'T's
    'Top T  ¯|¯
    
    For i = 1 To 7
        j = 16 - 2 * i 'width to copy
        
        BitBlt myDC, TILEW + i, TILEH - i, j, 1, myDC, i, TILEH - i, vbSrcCopy
    Next
    
    'Left T |-
    
    For i = 1 To 7
        j = 16 - 2 * i 'height to copy
        
        BitBlt myDC, TILEW - i, TILEH + i, 1, j, myDC, TILEW - i, i, vbSrcCopy
    Next
    
    'Bottom T _|_
    
    For i = 1 To 7
        j = 16 - 2 * i 'height to copy
        
        BitBlt myDC, TILEW + i, TILEH * 2 - 1 + i, j, 1, myDC, i, TILEH * 2 - 1 + i, vbSrcCopy
    Next
    
    'Right T  -|
    
    For i = 1 To 7
        j = 16 - 2 * i 'height to copy
        
        BitBlt myDC, TILEW * 2 - 1 + i, TILEH + i, 1, j, myDC, TILEW * 2 - 1 + i, TILEH * 2 + i, vbSrcCopy
    Next
    
    
    'Straights...
    
    'Vertical
    BitBlt myDC, TILEW * 3, 0, 8, TILEH * 3, myDC, 0, 0, vbSrcCopy
    BitBlt myDC, TILEW * 3 + 8, 0, 8, TILEH * 3, myDC, TILEW * 2 + 8, 0, vbSrcCopy
    
    'Horizontal
    BitBlt myDC, 0, TILEH * 3, TILEW * 3, 8, myDC, 0, 0, vbSrcCopy
    BitBlt myDC, 0, TILEH * 3 + 8, TILEW * 3, 8, myDC, 0, TILEH * 2 + 8, vbSrcCopy
    
    'Single wall...
    BitBlt myDC, TILEW * 3, TILEH * 3, 8, TILEH, myDC, 0, TILEH * 3, vbSrcCopy
    BitBlt myDC, TILEW * 3 + 8, TILEH * 3, 8, TILEH, myDC, TILEW * 2 + 8, TILEH * 3, vbSrcCopy
    
    picgenWallTile.Refresh
    
End Sub



Private Sub FlipV(ByRef srcPic As PictureBox, Left As Integer, width As Integer, Top As Integer, Height As Integer)
    Dim i As Integer
    
    pictemp.width = width
    pictemp.Height = Height
    
    BitBlt pictemp.hDc, 0, 0, pictemp.width, pictemp.Height, srcPic.hDc, Left, Top, vbSrcCopy
    pictemp.Refresh
    For i = 0 To Height - 1
        BitBlt srcPic.hDc, Left, Top + i, pictemp.width, 1, pictemp.hDc, 0, pictemp.Height - 1 - i, vbSrcCopy
    Next
    
    srcPic.Refresh
    pictemp.Cls
    
End Sub

Private Sub FlipH(ByRef srcPic As PictureBox, Left As Integer, width As Integer, Top As Integer, Height As Integer)
    Dim i As Integer
    
    pictemp.width = width
    pictemp.Height = Height
    
    BitBlt pictemp.hDc, 0, 0, pictemp.width, pictemp.Height, srcPic.hDc, Left, Top, vbSrcCopy
    pictemp.Refresh
    For i = 0 To width - 1
        BitBlt srcPic.hDc, Left + i, Top, 1, pictemp.Height, pictemp.hDc, pictemp.width - 1 - i, 0, vbSrcCopy
    Next
    
    srcPic.Refresh
    pictemp.Cls
    

End Sub
