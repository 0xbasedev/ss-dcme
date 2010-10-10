VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmGeneral 
   BackColor       =   &H8000000C&
   Caption         =   "Drake Continuum Map Editor"
   ClientHeight    =   11385
   ClientLeft      =   2685
   ClientTop       =   -1185
   ClientWidth     =   13860
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   WindowState     =   2  'Maximized
   Begin VB.PictureBox tlbTabs 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   220
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   13860
      TabIndex        =   171
      Top             =   840
      Width           =   13860
      Begin VB.CommandButton cmdChangeTabPos 
         Caption         =   "Ú"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   172
         TabStop         =   0   'False
         ToolTipText     =   "Change Tab Position"
         Top             =   0
         Width           =   375
      End
      Begin MSComctlLib.TabStrip tbMaps 
         CausesValidation=   0   'False
         Height          =   240
         Left            =   360
         TabIndex        =   173
         Top             =   0
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   423
         MultiRow        =   -1  'True
         TabFixedWidth   =   1849
         Placement       =   1
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Key             =   "Dummy"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   3000
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbarleft 
      Align           =   3  'Align Left
      Height          =   10320
      Left            =   0
      TabIndex        =   138
      Top             =   1065
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   18203
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Magnifier"
            Object.ToolTipText     =   "Magnifier (Z)"
            ImageKey        =   "Magnifier"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Selection"
            Object.ToolTipText     =   "Selection (S)"
            ImageKey        =   "Selection"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MagicWand"
            Object.ToolTipText     =   "Magic Wand (W)"
            ImageKey        =   "MagicWand"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FreehandSelection"
            Object.ToolTipText     =   "Freehand Selection"
            ImageKey        =   "Freehand"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Hand"
            Object.ToolTipText     =   "Hand (H)"
            ImageKey        =   "Hand"
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pencil"
            Object.ToolTipText     =   "Pencil (P)"
            ImageKey        =   "Pencil"
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dropper"
            Object.ToolTipText     =   "Dropper (D)"
            ImageKey        =   "Dropper"
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eraser"
            Object.ToolTipText     =   "Eraser (E)"
            ImageKey        =   "Eraser"
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Airbrush"
            Object.ToolTipText     =   "Airbrush (A)"
            ImageKey        =   "Airbrush"
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReplaceBrush"
            Object.ToolTipText     =   "Replace Brush (B)"
            ImageKey        =   "ReplaceBrush"
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bucket"
            Object.ToolTipText     =   "Bucket (F)"
            ImageKey        =   "Bucket"
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Line"
            Object.ToolTipText     =   "Line (L)"
            ImageKey        =   "Line"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SpLine"
            Object.ToolTipText     =   "Sticky Line (I)"
            ImageKey        =   "SpLine"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rectangle"
            Object.ToolTipText     =   "Rectangle (R)"
            ImageKey        =   "Rectangle"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ellipse"
            Object.ToolTipText     =   "Ellipse (O)"
            ImageKey        =   "Ellipse"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filled Rectangle"
            Object.ToolTipText     =   "Filled Rectangle"
            ImageKey        =   "FilledRectangle"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filled Ellipse"
            Object.ToolTipText     =   "Filled Ellipse"
            ImageKey        =   "FilledEllipse"
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CustomShape"
            Object.ToolTipText     =   "Other Shapes"
            ImageKey        =   "Cogwheel"
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileText"
            Object.ToolTipText     =   "Tile Text"
            ImageKey        =   "TileText"
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Regions"
            Object.ToolTipText     =   "Regions"
            ImageKey        =   "Regions"
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TestMap"
            Object.ToolTipText     =   "Test Map"
            ImageKey        =   "TestMap"
            Style           =   2
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LVZ"
            Object.ToolTipText     =   "LVZ Selector"
            ImageKey        =   "Package"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilImages 
      Left            =   720
      Top             =   1680
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
            Picture         =   "frmGeneral.frx":2C806
            Key             =   "Ellipse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2C91A
            Key             =   "NotRMode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2CC6E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2CD86
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2CE9E
            Key             =   "Region"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2D1F2
            Key             =   "Filled"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2D546
            Key             =   "Flip"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2D89A
            Key             =   "Mirror"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2DBEE
            Key             =   "NoClip"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2DF42
            Key             =   "Clip"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2E296
            Key             =   "R180"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2E5EA
            Key             =   "RLeft"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2E93E
            Key             =   "RRight"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2EC92
            Key             =   "UnFilled"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2EFE6
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F0FA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F20E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F322
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F436
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F54A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F65E
            Key             =   "Rectangle"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F772
            Key             =   "Arc"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2F886
            Key             =   "RMode"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2FBDA
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":2FF2E
            Key             =   "Segment"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":30282
            Key             =   "Zoom"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":305D6
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":3092A
            Key             =   "TextDef"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   600
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picRightBar 
      Align           =   4  'Align Right
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10320
      Left            =   9225
      ScaleHeight     =   684
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   139
      Top             =   1065
      Width           =   4630
      Begin DCME.cTileset cTileset 
         Height          =   3735
         Left            =   0
         TabIndex        =   195
         Top             =   525
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6588
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox pictilesetlarge 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   3180
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   92
         TabIndex        =   141
         Top             =   0
         Width           =   1380
         Begin VB.Label lblswitchtiles 
            BackStyle       =   0  'Transparent
            Caption         =   "<->"
            Height          =   255
            Left            =   555
            TabIndex        =   142
            Top             =   120
            Width           =   255
         End
         Begin VB.Shape shppreviewsel 
            BorderColor     =   &H000000FF&
            Height          =   510
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   510
         End
         Begin VB.Shape shppreviewsel 
            BorderColor     =   &H0000FFFF&
            Height          =   510
            Index           =   2
            Left            =   840
            Top             =   0
            Width           =   510
         End
      End
      Begin VB.PictureBox picdefaultwalltiles 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   3840
         Picture         =   "frmGeneral.frx":30C7E
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   183
         Top             =   8040
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.PictureBox picsmalltilepreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   0
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   143
         Top             =   360
         Visible         =   0   'False
         Width           =   270
         Begin VB.Shape shpsmalltilepreview 
            BorderColor     =   &H0000FFFF&
            Height          =   270
            Index           =   2
            Left            =   0
            Top             =   285
            Width           =   270
         End
         Begin VB.Shape shpsmalltilepreview 
            BorderColor     =   &H000000FF&
            Height          =   270
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   270
         End
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   2640
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   170
         Top             =   10800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   2520
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   169
         Top             =   10680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   2400
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   167
         Top             =   10560
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2280
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   166
         Top             =   10440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   2280
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   165
         Top             =   10320
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   2160
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   164
         Top             =   10200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   2040
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   162
         Top             =   10080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   56
         TabIndex        =   160
         Top             =   9840
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CheckBox chkForceFullRadar 
         Caption         =   "Full Map Preview"
         Height          =   255
         Left            =   2280
         TabIndex        =   159
         Top             =   9480
         Width           =   2175
      End
      Begin VB.PictureBox picDefaultShips 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         Picture         =   "frmGeneral.frx":33CC0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   148
         Top             =   5250
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picRadarPopup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   0
         Picture         =   "frmGeneral.frx":163902
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   155
         Top             =   5160
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox pichighlightspecial 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   3720
         Picture         =   "frmGeneral.frx":163E44
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   158
         Top             =   6840
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.PictureBox pictemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   151
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TileTextData 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   157
         Top             =   6600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.PictureBox picicons 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   0
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   145
         Top             =   4440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picdefaulttileset 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   3120
         Picture         =   "frmGeneral.frx":173A86
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   304
         TabIndex        =   168
         Top             =   9960
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.CommandButton cmdToggleRightPanel 
         Caption         =   "->"
         Height          =   330
         Left            =   0
         TabIndex        =   140
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picspecial 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   3840
         Picture         =   "frmGeneral.frx":17E188
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   156
         Top             =   5280
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.PictureBox piczoomtileset 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   2640
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   312
         TabIndex        =   144
         Top             =   1080
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.PictureBox picradar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   4320
         Left            =   120
         ScaleHeight     =   288
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   288
         TabIndex        =   154
         Top             =   5160
         Width           =   4320
      End
      Begin MSComctlLib.Toolbar tlbTileset 
         Height          =   330
         Left            =   360
         TabIndex        =   182
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditTileset"
               Object.ToolTipText     =   "Edit Tileset..."
               ImageIndex      =   77
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ImportTileset"
               Object.ToolTipText     =   "Import Tileset..."
               ImageIndex      =   61
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ExportTileset"
               Object.ToolTipText     =   "Export Tileset..."
               ImageIndex      =   62
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DiscardTileset"
               Object.ToolTipText     =   "Discard Tileset"
               ImageIndex      =   63
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditWalltiles"
               Object.ToolTipText     =   "Edit Walltiles..."
               ImageKey        =   "WallTiles"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditLVZ"
               Object.ToolTipText     =   "Manage LVZ..."
               ImageKey        =   "Package"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "[ debug -  label7 ]"
         Height          =   255
         Left            =   120
         TabIndex        =   163
         Top             =   10080
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "[ debug -  label6 ]"
         Height          =   255
         Left            =   120
         TabIndex        =   161
         Top             =   9840
         Width           =   4215
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   153
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label lblFrom 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   150
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label lblToA 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance:"
         Height          =   255
         Left            =   360
         TabIndex        =   152
         Top             =   4920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblFromA 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   255
         Left            =   360
         TabIndex        =   149
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblposition 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   147
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label lblpositionA 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         Height          =   255
         Left            =   360
         TabIndex        =   146
         Top             =   4440
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   304
         Y1              =   288
         Y2              =   288
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   65280
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   91
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18DDCA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18E364
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18E8FE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18EC98
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18F232
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18F5CC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18F966
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":18FF00
            Key             =   "PasteUnder"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19029A
            Key             =   "PasteNormal"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":190634
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":190BCE
            Key             =   "PasteTransparent"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":190F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":191502
            Key             =   "Mirror"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":191A9C
            Key             =   "Flip"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":192036
            Key             =   "Rotate"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1925D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":192B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":193104
            Key             =   "Selection"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19369E
            Key             =   "Magnifier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":193A38
            Key             =   "Dropper"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":193FD2
            Key             =   "Brush"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19456C
            Key             =   "Bucket"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":194906
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":194EA0
            Key             =   "Rectangle"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19543A
            Key             =   "FilledRectangle"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1959D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":195F6E
            Key             =   "Ellipse"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":196508
            Key             =   "FilledEllipse"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":196AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19703C
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1975D6
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":197970
            Key             =   "Eraser"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":197D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1982A4
            Key             =   "Grid"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19863E
            Key             =   "Tilenr"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1989D8
            Key             =   "Hand"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":198F72
            Key             =   "SpLine"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19930C
            Key             =   "FillInScreen"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1996A6
            Key             =   "TTM"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":199A40
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":199DDA
            Key             =   "PTM"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19A174
            Key             =   "Airbrush"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19A50E
            Key             =   "TilenrSelected"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19A8A8
            Key             =   "GridSelected"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19AC42
            Key             =   "FillInScreenSelected"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19AFDC
            Key             =   "FillInShapesSelected"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19B376
            Key             =   "Paste UnderSelected"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19B710
            Key             =   "Paste NormalSelected"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19BAAA
            Key             =   "Tip"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19BE44
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19C1DE
            Key             =   "Count"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19C578
            Key             =   "TileVertically"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19C912
            Key             =   "TileHorizontally"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19CCAC
            Key             =   "Cascade"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19D046
            Key             =   "MagicWand"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19D3E0
            Key             =   "WallTiles"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19D77A
            Key             =   "WallTilesSelected"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19DB14
            Key             =   "Paste TransparentSelected"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19DEAE
            Key             =   "ConvertToWalltiles"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19E248
            Key             =   "CheckForUpdates"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19E5E2
            Key             =   "ImportTileset"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19E97C
            Key             =   "ExportTileset"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19ED16
            Key             =   "DiscardTileset"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19F0B0
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19F44A
            Key             =   "SelectAdd"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19F7E4
            Key             =   "SelectRemove"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19FB7E
            Key             =   "Resize"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":19FF18
            Key             =   "CenterInScreen"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A02B2
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A064C
            Key             =   "SelectNone"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A09E6
            Key             =   "SelectLeftTile"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A0D80
            Key             =   "SelectRightTile"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A111A
            Key             =   "SelectAddLeftTile"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A14B4
            Key             =   "SelectAddRightTile"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A184E
            Key             =   "SelectRemoveLeftTile"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A1BE8
            Key             =   "SelectRemoveRightTile"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A1F82
            Key             =   "EditTileset"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A231C
            Key             =   "ZoomIn"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A26E4
            Key             =   "ZoomOut"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A2AA8
            Key             =   "LVZ"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A2E2F
            Key             =   "ReplaceBrush"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A31C9
            Key             =   "Cogwheel"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A3563
            Key             =   "TileText"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A38F7
            Key             =   "TestMap"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A3E91
            Key             =   "Regions"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A422B
            Key             =   "elvl"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A45C5
            Key             =   "EditTileText"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A4917
            Key             =   "PackageSelected"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A4CB1
            Key             =   "Package"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A504B
            Key             =   "RegionsSelected"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneral.frx":1A53E5
            Key             =   "Freehand"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolOptions 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   13860
      _ExtentX        =   24448
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlLarge"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Hand"
         Height          =   360
         Index           =   4
         Left            =   0
         TabIndex        =   188
         Top             =   800
         Width           =   8655
         Begin VB.Label LabelTool 
            Caption         =   "Hand                             You can switch to the hand anytime by holding down the middle mouse button"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   189
            Top             =   120
            Width           =   7575
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Freehand Selection"
         Height          =   360
         Index           =   3
         Left            =   0
         TabIndex        =   186
         Top             =   800
         Width           =   8655
         Begin VB.Label LabelTool 
            Caption         =   "Freehand Selection        Use 'Shift' to add tiles to selection and 'Ctrl' to remove tiles from it."
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   187
            Top             =   120
            Width           =   7575
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Selection"
         Height          =   360
         Index           =   1
         Left            =   0
         TabIndex        =   184
         Top             =   600
         Width           =   8655
         Begin VB.Label LabelTool 
            Caption         =   "Selection                        Use 'Shift' to add tiles to selection and 'Ctrl' to remove tiles from it."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   185
            Top             =   120
            Width           =   7575
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Test Map"
         Height          =   400
         Index           =   20
         Left            =   0
         TabIndex        =   119
         Top             =   500
         Width           =   13335
         Begin VB.CommandButton cmdBrowseSettings 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   8040
            TabIndex        =   134
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdStopTest 
            Caption         =   "Stop (Esc)"
            Height          =   255
            Left            =   6720
            TabIndex        =   133
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdStartTest 
            Caption         =   "Go!"
            Height          =   255
            Left            =   5640
            TabIndex        =   132
            Top             =   120
            Width           =   975
         End
         Begin VB.Frame Frame3 
            Height          =   410
            Left            =   1080
            TabIndex        =   120
            Top             =   -30
            Width           =   1695
            Begin VB.CheckBox chkTileCollision 
               Caption         =   "Tile Collision"
               Height          =   255
               Left            =   120
               TabIndex        =   121
               Top             =   120
               Value           =   1  'Checked
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ship"
            Height          =   410
            Left            =   2880
            TabIndex        =   123
            Top             =   -30
            Width           =   2535
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   7
               Left            =   2160
               TabIndex        =   131
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   6
               Left            =   1920
               TabIndex        =   130
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   5
               Left            =   1680
               TabIndex        =   129
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   4
               Left            =   1440
               TabIndex        =   128
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   3
               Left            =   1200
               TabIndex        =   127
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   126
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   125
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton optShip 
               Caption         =   "Option1"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   124
               Top             =   120
               Value           =   -1  'True
               Width           =   255
            End
         End
         Begin VB.Label lblCurrentSettings 
            Caption         =   "Settings: <DEFAULT>"
            Height          =   255
            Left            =   9000
            TabIndex        =   135
            Top             =   150
            Width           =   4335
         End
         Begin VB.Label LabelTool 
            Caption         =   "Test Map"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   122
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "LVZ"
         Height          =   400
         Index           =   21
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Width           =   14895
         Begin VB.CommandButton cmdLvzGoto 
            Caption         =   "Go To"
            Height          =   255
            Left            =   11400
            TabIndex        =   194
            Top             =   80
            Width           =   735
         End
         Begin VB.TextBox txtLvzY 
            Height          =   285
            Left            =   10680
            TabIndex        =   192
            Text            =   "0"
            Top             =   80
            Width           =   615
         End
         Begin VB.TextBox txtLvzX 
            Height          =   285
            Left            =   10080
            TabIndex        =   191
            Text            =   "00000"
            Top             =   80
            Width           =   615
         End
         Begin VB.TextBox txtLvzSnap 
            Height          =   285
            Left            =   8640
            MaxLength       =   3
            TabIndex        =   180
            Text            =   "1"
            Top             =   80
            Width           =   375
         End
         Begin VB.TextBox txtLvzObjectID 
            Height          =   285
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   179
            Top             =   80
            Width           =   615
         End
         Begin VB.ComboBox cmbLvzDisplayType 
            Height          =   315
            ItemData        =   "frmGeneral.frx":1A577F
            Left            =   6120
            List            =   "frmGeneral.frx":1A5795
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   60
            Width           =   1695
         End
         Begin VB.ComboBox cmbLvzLayerType 
            Height          =   315
            ItemData        =   "frmGeneral.frx":1A57DF
            Left            =   4320
            List            =   "frmGeneral.frx":1A57FB
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   60
            Width           =   1695
         End
         Begin VB.TextBox txtLVZDisplayTime 
            Height          =   285
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   175
            Text            =   "8888"
            Top             =   80
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Position:"
            Height          =   255
            Left            =   9240
            TabIndex        =   193
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Snap: (pixels)"
            Height          =   420
            Left            =   7920
            TabIndex        =   181
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblLvzID 
            Caption         =   "ID:"
            Height          =   255
            Left            =   1320
            TabIndex        =   178
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Display Time: (1/10th)"
            Height          =   425
            Left            =   2520
            TabIndex        =   174
            Top             =   0
            Width           =   975
         End
         Begin VB.Label LabelTool 
            Caption         =   "LVZ"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   137
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Region"
         Height          =   400
         Index           =   19
         Left            =   0
         TabIndex        =   114
         Top             =   600
         Width           =   11295
         Begin VB.CommandButton cmdHideAllRegions 
            Caption         =   "Hide All"
            Height          =   255
            Left            =   6480
            TabIndex        =   190
            Top             =   120
            Width           =   1215
         End
         Begin DCME.LayerList llRegionList 
            Height          =   285
            Left            =   2880
            TabIndex        =   117
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
            LinesShown      =   20
         End
         Begin VB.OptionButton optRegionSel 
            Height          =   300
            Index           =   0
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   60
            Width           =   300
         End
         Begin VB.OptionButton optRegionSel 
            Height          =   300
            Index           =   1
            Left            =   1920
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   60
            Width           =   300
         End
         Begin VB.Label LabelTool 
            Caption         =   "Regions"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   118
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Ellipse"
         Height          =   400
         Index           =   14
         Left            =   0
         TabIndex        =   75
         Top             =   600
         Width           =   9975
         Begin VB.Frame frmRender 
            Height          =   385
            Index           =   14
            Left            =   5640
            TabIndex        =   78
            Top             =   -30
            Width           =   2895
            Begin VB.CheckBox chkRenderAfter 
               Caption         =   "Render Width After Drawing Only"
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   79
               Top             =   120
               Width           =   2745
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   14
            Left            =   1305
            TabIndex        =   76
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin DCME.ToolProperty toolStep 
            Height          =   345
            Index           =   14
            Left            =   3480
            TabIndex        =   77
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Ellipse"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Custom Shape"
         Height          =   400
         Index           =   17
         Left            =   0
         TabIndex        =   94
         Top             =   600
         Width           =   13575
         Begin VB.Frame frmRender 
            Height          =   375
            Index           =   17
            Left            =   8640
            TabIndex        =   112
            Top             =   0
            Width           =   2895
            Begin VB.CheckBox chkRenderAfter 
               Caption         =   "Render Width After Drawing Only"
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   113
               Top             =   120
               Width           =   2745
            End
         End
         Begin VB.Frame frmCustomShape 
            BorderStyle     =   0  'None
            Caption         =   "Ellipse"
            Height          =   400
            Index           =   2
            Left            =   3960
            TabIndex        =   107
            Top             =   500
            Width           =   10095
            Begin VB.PictureBox piccustomshapePreview 
               AutoRedraw      =   -1  'True
               Height          =   400
               Index           =   2
               Left            =   1320
               ScaleHeight     =   23
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   23
               TabIndex        =   108
               Top             =   0
               Width           =   400
            End
            Begin DCME.ToolProperty customShapeSize 
               Height          =   345
               Index           =   2
               Left            =   1920
               TabIndex        =   109
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin DCME.ToolProperty customShapeTeethNumber 
               Height          =   345
               Index           =   2
               Left            =   3960
               TabIndex        =   110
               Top             =   0
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin VB.Label LabelCustomShape 
               Caption         =   "Regular Shape"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   111
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.Frame frmCustomShape 
            BorderStyle     =   0  'None
            Caption         =   "Ellipse"
            Height          =   400
            Index           =   1
            Left            =   0
            TabIndex        =   101
            Top             =   500
            Width           =   10575
            Begin VB.PictureBox piccustomshapePreview 
               AutoRedraw      =   -1  'True
               Height          =   400
               Index           =   1
               Left            =   1320
               ScaleHeight     =   23
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   23
               TabIndex        =   102
               Top             =   0
               Width           =   400
            End
            Begin DCME.ToolProperty customShapeSize 
               Height          =   345
               Index           =   1
               Left            =   1905
               TabIndex        =   103
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin DCME.ToolProperty customShapeTeethSize 
               Height          =   345
               Index           =   1
               Left            =   6240
               TabIndex        =   104
               Top             =   0
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin DCME.ToolProperty customShapeTeethNumber 
               Height          =   345
               Index           =   1
               Left            =   3960
               TabIndex        =   105
               Top             =   0
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin VB.Label LabelCustomShape 
               Caption         =   "Star"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   106
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.Frame frmCustomShape 
            BorderStyle     =   0  'None
            Caption         =   "Ellipse"
            Height          =   400
            Index           =   0
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   10575
            Begin VB.PictureBox piccustomshapePreview 
               AutoRedraw      =   -1  'True
               Height          =   400
               Index           =   0
               Left            =   1320
               ScaleHeight     =   23
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   23
               TabIndex        =   96
               Top             =   0
               Width           =   400
            End
            Begin DCME.ToolProperty customShapeSize 
               Height          =   345
               Index           =   0
               Left            =   1905
               TabIndex        =   97
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin DCME.ToolProperty customShapeTeethSize 
               Height          =   345
               Index           =   0
               Left            =   6240
               TabIndex        =   98
               Top             =   0
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin DCME.ToolProperty customShapeTeethNumber 
               Height          =   345
               Index           =   0
               Left            =   3960
               TabIndex        =   99
               Top             =   0
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   609
               Caption         =   ""
               Min             =   1
               Max             =   128
               Value           =   1
            End
            Begin VB.Label LabelCustomShape 
               Caption         =   "Cogwheel"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   100
               Top             =   120
               Width           =   1095
            End
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Eraser"
         Height          =   400
         Index           =   7
         Left            =   0
         TabIndex        =   22
         Top             =   600
         Width           =   6255
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   7
            Left            =   1320
            TabIndex        =   23
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   7
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5869
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   7
               Left            =   360
               Picture         =   "frmGeneral.frx":1A58DF
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   7
            Left            =   2160
            TabIndex        =   26
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Eraser"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Pencil"
         Height          =   400
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   600
         Width           =   10935
         Begin VB.CheckBox chkAdvancedPencil 
            Caption         =   "Ignore Special Tiles (Not Recommended)"
            Height          =   255
            Left            =   4200
            TabIndex        =   16
            ToolTipText     =   $"frmGeneral.frx":1A5955
            Top             =   75
            Width           =   3495
         End
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   5
            Left            =   1320
            TabIndex        =   12
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   5
               Left            =   360
               Picture         =   "frmGeneral.frx":1A5A09
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   5
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5A7F
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   5
            Left            =   2160
            TabIndex        =   15
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Pencil"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Filled Ellipse"
         Height          =   400
         Index           =   16
         Left            =   0
         TabIndex        =   88
         Top             =   600
         Width           =   9255
         Begin VB.Frame frmRender 
            Height          =   375
            Index           =   16
            Left            =   5640
            TabIndex        =   89
            Top             =   -30
            Width           =   2895
            Begin VB.CheckBox chkRenderAfter 
               Caption         =   "Render Width After Drawing Only"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   90
               Top             =   120
               Width           =   2745
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   16
            Left            =   1305
            TabIndex        =   91
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin DCME.ToolProperty toolStep 
            Height          =   345
            Index           =   16
            Left            =   3480
            TabIndex        =   92
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Filled Ellipse"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   93
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Filled Rectangle"
         Height          =   400
         Index           =   15
         Left            =   0
         TabIndex        =   81
         Top             =   600
         Width           =   6375
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   15
            Left            =   1320
            TabIndex        =   82
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   15
               Left            =   360
               Picture         =   "frmGeneral.frx":1A5AF5
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   15
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5B6B
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   15
            Left            =   2160
            TabIndex        =   85
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin DCME.ToolProperty toolStep 
            Height          =   345
            Index           =   15
            Left            =   4200
            TabIndex        =   86
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Filled Rectangle"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Rectangle"
         Height          =   400
         Index           =   13
         Left            =   0
         TabIndex        =   68
         Top             =   600
         Width           =   6375
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   13
            Left            =   1320
            TabIndex        =   69
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   13
               Left            =   360
               Picture         =   "frmGeneral.frx":1A5BE1
               Style           =   1  'Graphical
               TabIndex        =   71
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   13
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5C57
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   13
            Left            =   2160
            TabIndex        =   72
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin DCME.ToolProperty toolStep 
            Height          =   345
            Index           =   13
            Left            =   4200
            TabIndex        =   73
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Rectangle"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "ReplaceBrush"
         Height          =   400
         Index           =   12
         Left            =   0
         TabIndex        =   61
         Top             =   600
         Width           =   6375
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   12
            Left            =   1320
            TabIndex        =   62
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   12
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5CCD
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   12
               Left            =   360
               Picture         =   "frmGeneral.frx":1A5D43
               Style           =   1  'Graphical
               TabIndex        =   64
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   12
            Left            =   2160
            TabIndex        =   65
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin DCME.ToolProperty toolStep 
            Height          =   345
            Index           =   12
            Left            =   4200
            TabIndex        =   66
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Sticky Line"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Line"
         Height          =   400
         Index           =   11
         Left            =   0
         TabIndex        =   54
         Top             =   600
         Width           =   8415
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   11
            Left            =   1320
            TabIndex        =   55
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   11
               Left            =   360
               Picture         =   "frmGeneral.frx":1A5DB9
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   11
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5E2F
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   11
            Left            =   2160
            TabIndex        =   58
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin DCME.ToolProperty toolStep 
            Height          =   345
            Index           =   11
            Left            =   4200
            TabIndex        =   59
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Line"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   60
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Dropper"
         Height          =   400
         Index           =   6
         Left            =   0
         TabIndex        =   18
         Top             =   600
         Width           =   8655
         Begin VB.Frame Frame2 
            Height          =   390
            Left            =   1300
            TabIndex        =   20
            Top             =   -10
            Width           =   2055
            Begin VB.CheckBox chkDropperIgnoreEmpty 
               Caption         =   "Ignore Empty Tiles"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   135
               Width           =   1800
            End
         End
         Begin VB.Label LabelTool 
            Caption         =   "Dropper"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "ReplaceBrush"
         Height          =   400
         Index           =   9
         Left            =   0
         TabIndex        =   42
         Top             =   500
         Width           =   5775
         Begin VB.Frame frmToolTip 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   375
            Index           =   9
            Left            =   1320
            TabIndex        =   43
            Top             =   0
            Width           =   735
            Begin VB.OptionButton optToolRound 
               Height          =   300
               Index           =   9
               Left            =   0
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmGeneral.frx":1A5EA5
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   60
               Width           =   300
            End
            Begin VB.OptionButton optToolSquare 
               Height          =   300
               Index           =   9
               Left            =   360
               Picture         =   "frmGeneral.frx":1A5F1B
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   60
               Width           =   300
            End
         End
         Begin DCME.ToolProperty toolSize 
            Height          =   345
            Index           =   9
            Left            =   2160
            TabIndex        =   46
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Min             =   1
            Max             =   128
            Value           =   1
         End
         Begin VB.Label LabelTool 
            Caption         =   "Replace Brush"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "MagicWand"
         Height          =   400
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   8655
         Begin VB.Frame frmMagicWandScreen 
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   0
            Width           =   1815
            Begin VB.CheckBox chkMagicWandScreen 
               Caption         =   "Limit To Screen"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Value           =   1  'Checked
               Width           =   1560
            End
         End
         Begin VB.Frame Frame1 
            Height          =   375
            Left            =   1300
            TabIndex        =   6
            Top             =   0
            Width           =   1935
            Begin VB.CheckBox chkMagicWandDiagonal 
               Caption         =   "Check For Diagonals"
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   120
               Width           =   1785
            End
         End
         Begin VB.Label LabelTool 
            Caption         =   "Magic Wand"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Airbrush"
         Height          =   360
         Index           =   8
         Left            =   0
         TabIndex        =   28
         Top             =   600
         Width           =   11295
         Begin VB.Frame frmSize 
            Caption         =   "Radius"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8520
            TabIndex        =   38
            Top             =   -10
            Width           =   2055
            Begin VB.TextBox txtAirBrSize 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   720
               TabIndex        =   39
               Text            =   "999"
               Top             =   135
               Width           =   375
            End
            Begin MSComctlLib.Slider sldAirbSize 
               Height          =   210
               Left            =   1080
               TabIndex        =   40
               Top             =   135
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   370
               _Version        =   393216
               LargeChange     =   10
               Min             =   3
               Max             =   50
               SelStart        =   10
               TickStyle       =   3
               Value           =   10
            End
         End
         Begin VB.Frame frmDensity 
            Caption         =   "Density"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6240
            TabIndex        =   34
            Top             =   -10
            Width           =   2175
            Begin VB.TextBox txtAirBrDensity 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   630
               TabIndex        =   35
               Text            =   "999"
               Top             =   135
               Width           =   360
            End
            Begin MSComctlLib.Slider sldAirbDensity 
               Height          =   210
               Left            =   1200
               TabIndex        =   37
               Top             =   135
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   370
               _Version        =   393216
               LargeChange     =   10
               Min             =   1
               Max             =   200
               SelStart        =   1
               TickStyle       =   3
               Value           =   1
            End
            Begin VB.Label lblAirbDensity 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1065
               TabIndex        =   36
               Top             =   120
               Width           =   150
            End
         End
         Begin VB.Frame frmAsteroids 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1300
            TabIndex        =   29
            Top             =   -10
            Width           =   4815
            Begin VB.CheckBox chkUseBigAsteroids 
               Caption         =   "Big"
               Height          =   195
               Left            =   4080
               TabIndex        =   31
               Top             =   120
               Width           =   615
            End
            Begin VB.CheckBox chkUseSmallAsteroids2 
               Caption         =   "Small2"
               Height          =   195
               Left            =   3120
               TabIndex        =   33
               Top             =   120
               Width           =   975
            End
            Begin VB.CheckBox chkuseSmallAsteroids1 
               Caption         =   "Small1"
               Height          =   195
               Left            =   2280
               TabIndex        =   32
               Top             =   120
               Width           =   975
            End
            Begin VB.CheckBox chkUseAsAsteroidBrush 
               Caption         =   "Use as Asteroid Brush"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   120
               Width           =   1935
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00C0C0C0&
               X1              =   3000
               X2              =   3000
               Y1              =   90
               Y2              =   360
            End
         End
         Begin VB.Label LabelTool 
            Caption         =   "Airbrush"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Magnifier"
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   8655
         Begin VB.Label LabelTool 
            Caption         =   "Magnifier"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lblcurrentzoom 
            Caption         =   "Zoom:"
            Height          =   255
            Left            =   1300
            TabIndex        =   4
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmTool 
         BorderStyle     =   0  'None
         Caption         =   "Bucket"
         Height          =   400
         Index           =   10
         Left            =   0
         TabIndex        =   48
         Top             =   600
         Width           =   8655
         Begin VB.Frame frmBucketFillInScreen 
            Height          =   390
            Left            =   3000
            TabIndex        =   52
            Top             =   0
            Width           =   2055
            Begin VB.CheckBox chkFillInScreen 
               Caption         =   "Limit Fill To Screen"
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   135
               Value           =   1  'Checked
               Width           =   1800
            End
         End
         Begin VB.Frame frmBucket 
            Height          =   390
            Left            =   1300
            TabIndex        =   50
            Top             =   0
            Width           =   1575
            Begin VB.CheckBox chkFillDiagonal 
               Caption         =   "Fill Diagonally"
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   135
               Width           =   1320
            End
         End
         Begin VB.Label LabelTool 
            Caption         =   "Bucket"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.Toolbar toolbartop 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13860
      _ExtentX        =   24448
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   31
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New map"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open a map"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save map"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut the selection"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy the selection"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste the selection"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grid"
            Object.ToolTipText     =   "Show Grid"
            ImageKey        =   "Grid"
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileNr"
            Object.ToolTipText     =   "Show Tile Numbers"
            ImageKey        =   "Tilenr"
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowRegions"
            Object.ToolTipText     =   "Show Regions"
            ImageKey        =   "Regions"
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowLVZ"
            Object.ToolTipText     =   "Show LVZ"
            ImageKey        =   "Package"
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZoomIn"
            Object.ToolTipText     =   "Zoom In"
            ImageKey        =   "ZoomIn"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZoomOut"
            Object.ToolTipText     =   "Zoom Out"
            ImageKey        =   "ZoomOut"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mirror"
            Object.ToolTipText     =   "Flip Selection Horizontally"
            ImageKey        =   "Mirror"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Flip"
            Object.ToolTipText     =   "Flip Selection Vertically"
            ImageKey        =   "Flip"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rotate"
            Object.ToolTipText     =   "Rotate Selection..."
            ImageKey        =   "Rotate"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Replace"
            Object.ToolTipText     =   "Switch/Replace Tiles..."
            ImageKey        =   "Replace"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TextToMap"
            Object.ToolTipText     =   "Text To Map..."
            ImageKey        =   "TTM"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PicToMap"
            Object.ToolTipText     =   "Picture To Map..."
            ImageKey        =   "PTM"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PasteType"
            ImageKey        =   "PasteNormal"
            Style           =   5
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditELVL"
            Object.ToolTipText     =   "Edit eLVL Attributes..."
            ImageKey        =   "elvl"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnusep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save map and lvz's"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSaveSelect 
         Caption         =   "Sa&ve items..."
      End
      Begin VB.Menu mnuSaveMiniMap 
         Caption         =   "Save &Full Radar View..."
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportTileset 
         Caption         =   "&Import Tileset..."
      End
      Begin VB.Menu mnuExportTileset 
         Caption         =   "&Export Tileset..."
      End
      Begin VB.Menu mnudiscardtileset 
         Caption         =   "&Discard Tileset"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRevert 
         Caption         =   "Re&vert..."
      End
      Begin VB.Menu mnuOpenAutosave 
         Caption         =   "Open Autosave"
         Begin VB.Menu mnulstAutosaves 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "&Recent"
         Begin VB.Menu mnulstRecent 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Switch/Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuMirror 
         Caption         =   "Flip &Horizontal"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "Flip &Vertical"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "R&otate..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Resi&ze..."
      End
      Begin VB.Menu mnusep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWallTiles 
         Caption         =   "Edit &Walltiles..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditTileset 
         Caption         =   "Edit Ti&leset..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTileText 
         Caption         =   "Edit Tile-Text..."
      End
      Begin VB.Menu mnusep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCount 
         Caption         =   "Cou&nt Tiles..."
      End
      Begin VB.Menu mnuTextToMap 
         Caption         =   "Add Te&xt To Map..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuPTM 
         Caption         =   "Add P&icture To Map..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnusep26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuElvl 
         Caption         =   "eLVL attributes..."
      End
      Begin VB.Menu mnuManageLVZ 
         Caption         =   "Manage LVZ..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuGrid 
         Caption         =   "Show &Grid"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuTileNR 
         Caption         =   "Show T&ile Numbers"
      End
      Begin VB.Menu mnuShowRegions 
         Caption         =   "Show Regions"
      End
      Begin VB.Menu mnuShowLVZ 
         Caption         =   "Show LVZ"
      End
      Begin VB.Menu mnusep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawMode 
         Caption         =   "Draw Mode"
         Begin VB.Menu mnuNormalPaste 
            Caption         =   "&Normal Draw"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPasteUnder 
            Caption         =   "Draw &Under Existing Tiles"
         End
         Begin VB.Menu mnuTransparentPaste 
            Caption         =   "&Transparent Selection"
         End
      End
      Begin VB.Menu mnusep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences..."
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "&Selection"
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select"
         Begin VB.Menu mnuSelectAllMap 
            Caption         =   "Select A&ll"
         End
         Begin VB.Menu mnuSelectAll 
            Caption         =   "Select &All Tiles (visible only)"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnusep18 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSelectLeftSelectedTileMap 
            Caption         =   "Left Selected Tile"
         End
         Begin VB.Menu mnuSelectLeftSelectedTile 
            Caption         =   "&Left Selected Tile (visible only)"
         End
         Begin VB.Menu mnuSelectRightSelectedTileMap 
            Caption         =   "Right Selected Tile"
         End
         Begin VB.Menu mnuSelectRightSelectedTile 
            Caption         =   "&Right Selected Tile (visible only)"
         End
      End
      Begin VB.Menu mnuAddToSelection 
         Caption         =   "A&dd To Selection"
         Begin VB.Menu mnuAddAll 
            Caption         =   "&Add All (visible only)"
         End
         Begin VB.Menu mnusep20 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddLeftSelectedTileMap 
            Caption         =   "Left Selected Tile"
         End
         Begin VB.Menu mnuAddLeftSelectedTile 
            Caption         =   "&Left Selected Tile (visible only)"
         End
         Begin VB.Menu mnuAddRightSelectedTileMap 
            Caption         =   "Right Selected Tile"
         End
         Begin VB.Menu mnuAddRightSelectedTile 
            Caption         =   "&Right Selected Tile (visible only)"
         End
      End
      Begin VB.Menu mnuRemoveFromSelection 
         Caption         =   "&Remove From Selection"
         Begin VB.Menu mnuSelectNone 
            Caption         =   "&Deselect All"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuSelectNoneVisible 
            Caption         =   "Deselect All (&visible only)"
         End
         Begin VB.Menu mnusep19 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRemoveLeftSelectedTileMap 
            Caption         =   "Left Selected Tile"
         End
         Begin VB.Menu mnuRemoveLeftSelectedTile 
            Caption         =   "&Left Selected Tile (in screen)"
         End
         Begin VB.Menu mnuRemoveRightSelectedTileMap 
            Caption         =   "Right Selected Tile"
         End
         Begin VB.Menu mnuRemoveRightSelectedTile 
            Caption         =   "&Right Selected Tile (in screen)"
         End
      End
      Begin VB.Menu mnusep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvtowalltiles 
         Caption         =   "Convert Selection To &Walltiles"
      End
      Begin VB.Menu mnusep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterSelectionMap 
         Caption         =   "Center Selection on &Map"
      End
      Begin VB.Menu mnuCenterSelection 
         Caption         =   "&Center Selection on Screen"
      End
      Begin VB.Menu mnusep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteSelection 
         Caption         =   "Delete Selection"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuBookmark 
         Caption         =   "Goto"
         Index           =   0
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnusep24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetBookmarks 
         Caption         =   "&Set Bookmark"
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#0 (Shift+0)"
            Index           =   0
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#1 (Shift+1)"
            Index           =   1
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#2 (Shift+2)"
            Index           =   2
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#3 (Shift+3)"
            Index           =   3
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#4 (Shift+4)"
            Index           =   4
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#5 (Shift+5)"
            Index           =   5
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#6 (Shift+6)"
            Index           =   6
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#7 (Shift+7)"
            Index           =   7
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#8 (Shift+8)"
            Index           =   8
         End
         Begin VB.Menu mnuSetBookmark 
            Caption         =   "#9 (Shift+9)"
            Index           =   9
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnusep23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbars 
         Caption         =   "Toolbars"
         Begin VB.Menu mnuToolbarStandard 
            Caption         =   "Standard"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuToolbarTools 
            Caption         =   "Tools"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuToolbarToolOptions 
            Caption         =   "Tool Options"
         End
         Begin VB.Menu mnuToolbarMapTabs 
            Caption         =   "Map Tabs"
         End
         Begin VB.Menu mnusep25 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTogglePinToolOptions 
            Caption         =   "Pin Tool Options"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaps 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTips 
         Caption         =   "Show &Tips..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpdates 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu mnusep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugLog 
         Caption         =   "View &Debug Log..."
      End
      Begin VB.Menu mnuShowDebugInfo 
         Caption         =   "Debug Mode"
      End
      Begin VB.Menu mnusep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About DCME..."
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Begin VB.Menu mnuSpeedTest 
         Caption         =   "Speed Test"
      End
      Begin VB.Menu mnuShowMemUsage 
         Caption         =   "Memory Usage..."
      End
      Begin VB.Menu mnuDebugLayers 
         Caption         =   "Layer Screenshots"
      End
   End
   Begin VB.Menu customShapeMenu 
      Caption         =   "_CustomShapeMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCustomShape 
         Caption         =   "Cogwheel"
         Index           =   0
      End
      Begin VB.Menu mnuCustomShape 
         Caption         =   "Star"
         Index           =   1
      End
      Begin VB.Menu mnuCustomShape 
         Caption         =   "Regular Shape"
         Index           =   2
      End
   End
   Begin VB.Menu mnuLvz 
      Caption         =   "_LvzMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuLvzAddImage 
         Caption         =   "&Add Image..."
      End
      Begin VB.Menu mnuLvzEditImage 
         Caption         =   "&Edit Image..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLvzEditAnimation 
         Caption         =   "Edit A&nimation Settings..."
      End
      Begin VB.Menu mnuLvzDeleteImage 
         Caption         =   "&Delete Image"
      End
      Begin VB.Menu mnuLvzOpenManager 
         Caption         =   "Open LVZ &Manager..."
      End
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'holds whether map i (0..9) is loaded
Dim loadedmaps(9) As Boolean

Dim tabloaded(9) As Boolean


'holds the actual references to the maps
Dim Maps(9) As frmMain


'Display layers of active map
Dim MapLayers(DL_Regions To DL_Buffer) As clsDisplayLayer


Public clipboard As New ClipB


'index of the current active map
Public activemap As Integer

Dim FloatTileset As New frmFloatTileset
Public FloatRadar As New frmFloatRadar

Public AutoHideTileset As Boolean
Public AutoHideRadar As Boolean







'used to avoid resizing radar when hiding right panel
Dim dontResizeRadar As Boolean
'used to avoid updating preview when hiding/showing popup radar or tileset
Public dontUpdatePreview As Boolean

'checks if the update is triggered automaticly
Public quickupdate As Boolean

'is the update form already loaded
Public updateformloaded As Boolean

Public updateready As Boolean
Public updatefilepath As String



Dim lastwindowstate As FormWindowStateConstants



Dim c_isbusy As Boolean

Dim c_showprogress As Boolean
Dim c_progress As Long

Dim c_toolOptionExists(1 To TOOLCOUNT) As Boolean

Dim dontRefreshRegions As Boolean





'Tileset color settings
Private m_lLeftTilesetColor As Long
Private m_lRightTilesetColor As Long
Private m_lTilesetBackgroundColor As Long 'Color behind the tile preview's








Private Function SpeedTest() As Double
    If Not loadedmaps(activemap) Then Exit Function
    
      'Test sub to test various elements their speed
          Dim lbx As Integer    'lowestval for x
          Dim hbx As Integer    'highestval for x
          Dim lby As Integer    'lowestval for y
          Dim hby As Integer    'highestval for y
        lbx = (Maps(activemap).hScr.value \ (TILEW * Maps(activemap).magnifier.zoom))
        hbx = ((Maps(activemap).hScr.value + Maps(activemap).picPreview.width) \ (TILEW * Maps(activemap).magnifier.zoom))
        lby = (Maps(activemap).vScr.value \ (TILEW * Maps(activemap).magnifier.zoom))
        hby = ((Maps(activemap).vScr.value + Maps(activemap).picPreview.height) \ (TILEW * Maps(activemap).magnifier.zoom))

          Dim tick As Long
          Dim finalresult As Long
        finalresult = 0

          Dim i As Long
          Dim j As Integer
' LVZ interface test
'            Const nrOfTests = 500
'
'            tick = GetTickCount
'
'            For i = 0 To nrOfTests - 1
'
'                Call Maps(activemap).tileset.TickAnimation
'
'            Next
          


'    'UPDATELEVEL TEST
          Const nrOfTests = 100

        Dim ticktotal() As Long
        ReDim ticktotal(DL_Regions To DL_Buffer + 1)

        Dim ticktest() As Long
        ReDim ticktest(DL_Regions To DL_Buffer + 1)


        tick = GetTickCount

        For i = 0 To nrOfTests - 1
              'INSERT METHOD HERE


            Call Maps(activemap).UpdateLevelTest(ticktest)
              '//////////////////

            For j = DL_Regions To DL_Buffer + 1
                ticktotal(j) = ticktotal(j) + ticktest(j)
            Next
       Next


' intersect test
'          Const nrOfTests = 10000000
'        Dim r1 As RECT, r2 As RECT, rr As RECT
'        r1.left = 500
'        r1.right = 600
'        r1.top = 500
'        r1.bottom = 600
'
'        r2.left = 550
'        r2.right = 800
'        r2.top = 400
'        r2.bottom = 520
'
'        For i = 0 To nrOfTests - 1
'              'INSERT METHOD HERE
'
'            rr = Intersection2(r1, r2)
'
'              '//////////////////
'
'       Next


    'redraw regions TEST
'          Const nrOfTests = 1
               
        
'        tick = GetTickCount
'
'        For i = 0 To nrOfTests - 1
'              'INSERT METHOD HERE
'
'
'            Call Maps(activemap).Regions.RedrawAllRegions
'              '//////////////////
'
'       Next
       
       


       SpeedTest = (GetTickCount - tick) / nrOfTests
       
       Dim result As String

       Dim ticktesttotal As Double

        For j = DL_Regions To DL_Buffer + 1
            ticktesttotal = ticktesttotal + (ticktotal(j) / nrOfTests)
            Select Case j
                Case DL_Regions
                    result = result & "Regions  "
                Case DL_LVZunder
                    result = result & "LVZ under"
                Case DL_Tiles
                    result = result & "Tiles        "
                Case DL_Selection
                    result = result & "Selection"
                Case DL_LVZover
                    result = result & "LVZ over "
                Case DL_Buffer
                    result = result & "Buffer   "
                Case Else
                    result = result & "Preview  "
            End Select
            result = result & ":  " & vbTab & ticktotal(j) / nrOfTests & "ms" & vbCrLf
        Next
        result = result & "Other         :  " & vbTab & SpeedTest - ticktesttotal & "ms" & vbCrLf
        result = result & vbCrLf
        result = result & "TOTAL    :  " & vbTab & ticktesttotal & "ms"
        MessageBox result
'            MessageBox SpeedTest
End Function



Private Sub chkAdvancedPencil_Click()
    If chkAdvancedPencil.value = vbChecked Then
        Call SetSetting("AdvancedPencil", "1")
    Else
        Call SetSetting("AdvancedPencil", "0")
    End If
End Sub

Private Sub chkDropperIgnoreEmpty_Click()
    If chkDropperIgnoreEmpty.value = vbChecked Then
        Call SetSetting("DropEmpty", "1")
    Else
        Call SetSetting("DropEmpty", "0")
    End If
End Sub



Private Sub chkFillDiagonal_Click()
    If chkFillDiagonal.value = vbChecked Then
        Call SetSetting("FillDiagonal", "1")
    Else
        Call SetSetting("FillDiagonal", "0")
    End If
End Sub


Private Sub chkFillInScreen_Click()
    If chkFillInScreen.value = vbChecked Then
        Call SetSetting("FillInScreen", "1")
    Else
        Call SetSetting("FillInScreen", "0")
    End If
End Sub

Private Sub chkForceFullRadar_Click()
    Call SetSetting("ForceFullRadar", chkForceFullRadar.value)
    
    If loadedmaps(activemap) Then
        Call Maps(activemap).UpdatePreview(False, True)
    End If
End Sub

'Private Sub chklvzforcetransparency_Click()
'10        Call SetSetting("ForceLVZTransparency", chkLvzForceTransparency.value)
'
'20        If loadedmaps(activemap) Then
'30            Call Maps(activemap).UpdatePreview(True, False)
'40        End If
'
'End Sub

Private Sub chkMagicWandDiagonal_Click()
    If chkMagicWandDiagonal.value = vbChecked Then
        Call SetSetting("WandDiagonal", "1")
    Else
        Call SetSetting("WandDiagonal", "0")
    End If
End Sub

Private Sub chkMagicWandScreen_Click()
    If chkMagicWandScreen.value = vbChecked Then
        Call SetSetting("WandScreen", "1")
    Else
        Call SetSetting("WandScreen", "0")
    End If
End Sub

Private Sub chkRenderAfter_Click(Index As Integer)
    If chkRenderAfter(curtool - 1).value = vbChecked Then
        Call SetSetting("ToolRenderAfter" & ToolName(curtool - 1), "1")
    Else
        Call SetSetting("ToolRenderAfter" & ToolName(curtool - 1), "0")
    End If
End Sub

'Private Sub chklvzsnaptotiles_Click()
'10        If chkLvzSnapToTiles.value = vbChecked Then
'20            Call SetSetting("SnapToTiles", "1")
'30        Else
'40            Call SetSetting("SnapToTiles", "0")
'50        End If
'End Sub

Private Sub chkTileCollision_GotFocus()
    If Maps(activemap).TestMap.isRunning Then

        If chkTileCollision.value = checked Then
            chkTileCollision.value = Unchecked
        Else
            chkTileCollision.value = checked
        End If

        If chkTileCollision.value = vbChecked Then
            Call SetSetting("TileCollision", "1")
        Else
            Call SetSetting("TileCollision", "0")
        End If

        Maps(activemap).picPreview.setfocus
    End If
End Sub

Private Sub chkUseAsAsteroidBrush_Click()
    If chkUseAsAsteroidBrush.value = vbChecked Then
        Call SetSetting("UseAirBrushAsAsteroids", "1")
        chkUseBigAsteroids.visible = True
        chkuseSmallAsteroids1.visible = True
        chkUseSmallAsteroids2.visible = True
        frmAsteroids.width = 4815
    Else
        Call SetSetting("UseAirBrushAsAsteroids", "0")
        chkUseBigAsteroids.visible = False
        chkuseSmallAsteroids1.visible = False
        chkUseSmallAsteroids2.visible = False
        frmAsteroids.width = 2175
    End If
    frmDensity.Left = frmAsteroids.Left + frmAsteroids.width + 125
    frmSize.Left = frmDensity.Left + frmDensity.width + 125
End Sub

Private Sub chkUseBigAsteroids_Click()
    If chkUseBigAsteroids.value = vbChecked Then
        Call SetSetting("UseBigAsteroids", "1")
    Else
        Call SetSetting("UseBigAsteroids", "0")
    End If
End Sub

Private Sub chkuseSmallAsteroids1_Click()
    If chkuseSmallAsteroids1.value = vbChecked Then
        Call SetSetting("UseSmallAsteroids1", "1")
    Else
        Call SetSetting("UseSmallAsteroids1", "0")
    End If
End Sub

Private Sub chkUseSmallAsteroids2_Click()
    If chkUseSmallAsteroids2.value = vbChecked Then
        Call SetSetting("UseSmallAsteroids2", "1")
    Else
        Call SetSetting("UseSmallAsteroids2", "0")
    End If
End Sub






Private Sub cmbLvzDisplayType_Click()
    If loadedmaps(activemap) And cmbLvzDisplayType.ListIndex <> -1 Then
        Call Maps(activemap).lvz.ChangeSelectionDisplayMode(cmbLvzDisplayType.ListIndex)
    End If
End Sub

Private Sub cmbLvzLayerType_Click()
    If loadedmaps(activemap) And cmbLvzLayerType.ListIndex <> -1 Then
        Call Maps(activemap).lvz.ChangeSelectionLayer(cmbLvzLayerType.ListIndex)

    End If
End Sub



'TOREMOVE---
'Private Sub cmbLVZTilesetDisplayType_Click()
'    If loadedmaps(activemap) Then
'        Maps(activemap).lvz.MapObjectDefaultMode = cmbLVZTilesetDisplayType.ListIndex
'    End If
'End Sub
'
'Private Sub cmbLVZTilesetLayerType_Click()
'    If loadedmaps(activemap) Then
'        Maps(activemap).lvz.MapObjectDefaultLayer = cmbLVZTilesetLayerType.ListIndex
'    End If
'End Sub

Private Sub cmdChangeTabPos_Click()
    If tlbTabs.Align = vbAlignBottom Then
        tlbTabs.Align = vbAlignTop
        tbMaps.Placement = tabPlacementTop
        cmdChangeTabPos.Caption = "Ú"
        
        Call SetSetting("MapTabPosition", tlbTabs.Align)
    Else
        tlbTabs.Align = vbAlignBottom
        tbMaps.Placement = tabPlacementBottom
        cmdChangeTabPos.Caption = "Ù"
        
        Call SetSetting("MapTabPosition", tlbTabs.Align)
    End If
    
End Sub

'Private Sub cmdLVZJumpTo_Click()
'    Dim objIdx As Integer
'    objIdx = val(Mid$(cmbMapObjects.Text, 14, Len(cmbMapObjects.Text) - 13))
'
'    'retrieve lvz owner of map object
'    Dim i As Integer
'    Dim lvzidx As Integer
'    lvzidx = -1
'    For i = cmbMapObjects.ListIndex To 0 Step -1
'        If Mid$(cmbMapObjects.list(i), 1, 4) <> "    " Then
'            lvzidx = Maps(activemap).lvz.getIndexOfLVZ(cmbMapObjects.list(i))
'            If lvzidx <> -1 Then
'                Exit For
'            End If
'        End If
'    Next
'
'    If lvzidx <> -1 Then
'        'retrieve coordinates of map object
'        Dim X As Integer
'        Dim Y As Integer
'        X = Maps(activemap).lvz.getLVZ(lvzidx).mapobjects(objIdx).X
'        Y = Maps(activemap).lvz.getLVZ(lvzidx).mapobjects(objIdx).Y
'
'        'move screen lbx & lby to those coordinates
'        Call Maps(activemap).SetScrollbarValues(X * (Maps(activemap).currenttilew / TileW), Y * (Maps(activemap).currenttilew / TileW), False)
'        Call Maps(activemap).UpdateLevel
'    End If
'End Sub

'Private Sub cmbMapObjects_Click()
'    If Not loadedmaps(activemap) Then Exit Sub
'
'    If Mid$(cmbMapObjects.Text, 1, 4) <> "    " Then
'        'we didn't have a map object selected, take the 1st following map object
'        Dim i As Integer
'        For i = cmbMapObjects.ListIndex To cmbMapObjects.ListCount - 1
'            If Mid$(cmbMapObjects.list(i), 1, 4) = "    " Then
'                cmbMapObjects.ListIndex = i
'                Exit Sub
'            End If
'        Next
'
'        'no mapobject found. There is one, else we would have been disabled.
'        'start looking backwards
'        For i = cmbMapObjects.ListIndex To 0 Step -1
'            If Mid$(cmbMapObjects.list(i), 1, 4) = "    " Then
'                cmbMapObjects.ListIndex = i
'                Exit Sub
'            End If
'        Next
'
'        'no map object found. ERROR, there should been at least listed one
'        messagebox "Error, No map objects listed ; cmbMapObjects_Click", vbOKOnly + vbExclamation
'    End If
'End Sub

Private Sub cmdBrowseSettings_Click()
    On Error GoTo errh
    
    If Not loadedmaps(activemap) Then Exit Sub
    
    cd.DialogTitle = "Select a settings file"
    cd.flags = cdlOFNHideReadOnly
    cd.Filter = "Settings files (*.cfg)|*.cfg"
    cd.ShowOpen
    
    If FileExists(cd.filename) Or Maps(activemap).CFG.GetCfgPath = "" Then
        Call Maps(activemap).CFG.SetCfgPath(cd.filename)
    End If
    
    If Maps(activemap).CFG.GetCfgPath <> "" Then
        
        If Maps(activemap).TestMap.isRunning Then
            Call Maps(activemap).TestMap.ReadSettings
        End If
    End If
    
    On Error Resume Next
    Maps(activemap).picPreview.setfocus
    
    Exit Sub

errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
        MessageBox Err & ":" & Err.description, vbCritical
    End If
End Sub

Private Sub cmdHideAllRegions_Click()
    If loadedmaps(activemap) Then
        Dim selregn As Integer
        selregn = llRegionList.ListIndex
        Call Maps(activemap).Regions.HideAllRegions
'        Call UpdateToolToolbar
        Call UpdateRegionList
        llRegionList.ListIndex = selregn
        llRegionList.HideList
    End If
    
End Sub

Private Sub cmdLvzGoto_Click()
    If loadedmaps(activemap) Then
        With Maps(activemap)
            Call .SetFocusAt(IIf(txtLvzX.Enabled, CInt(txtLvzX.Text) \ TILEW, .ScreenToTileX(.picPreview.width \ 2)), _
                        IIf(txtLvzY.Enabled, CInt(txtLvzY.Text) \ TILEW, .ScreenToTileY(.picPreview.height \ 2)), _
                        .picPreview.width \ 2, .picPreview.height \ 2, True)
        End With
    End If
End Sub

'Private Sub cmdLVZMoveToScreen_Click()
'    If Not loadedmaps(activemap) Then Exit Sub
'
'    Dim lbx As Integer    'lowestval for x
'    Dim lby As Integer    'lowestval for y
'    lbx = (Maps(activemap).Hscr.Value \ (TileW * Maps(activemap).magnifier.zoom))
'    lby = (Maps(activemap).Vscr.Value \ (TileW * Maps(activemap).magnifier.zoom))
'
'    Dim objIdx As Integer
'    objIdx = val(Mid$(cmbMapObjects.Text, 14, Len(cmbMapObjects.Text) - 13))
'
'    'retrieve lvz owner of map object
'    Dim i As Integer
'    Dim lvzidx As Integer
'    lvzidx = -1
'    For i = cmbMapObjects.ListIndex To 0 Step -1
'        If Mid$(cmbMapObjects.list(i), 1, 4) <> "    " Then
'            lvzidx = Maps(activemap).lvz.getIndexOfLVZ(cmbMapObjects.list(i))
'            If lvzidx <> -1 Then
'                Exit For
'            End If
'        End If
'    Next
'
'    If lvzidx <> -1 Then
'        'retrieve coordinates of map object
'        Dim X As Integer
'        Dim Y As Integer
'        Dim tmplvz As LVZstruct
'        tmplvz = Maps(activemap).lvz.getLVZ(lvzidx)
'
'        tmplvz.mapobjects(objIdx).X = lbx * TileW 'Maps(activemap).currenttilew
'        tmplvz.mapobjects(objIdx).Y = lby * TileW 'Maps(activemap).currenttilew
'
'        Call Maps(activemap).lvz.setLVZ(tmplvz, lvzidx)
'
'        'move screen lbx & lby to those coordinates
'        Call Maps(activemap).UpdatePreview
'    End If
'End Sub

Private Sub cmdStartTest_Click()
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).TestMap.isRunning Then
        Call Maps(activemap).TestMap.WarpShip
    Else
        Call Maps(activemap).TestMap.StartRun
    End If

    On Error Resume Next
    Maps(activemap).picPreview.setfocus
End Sub

Private Sub cmdStopTest_Click()
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).TestMap.isRunning Then
        Call Maps(activemap).TestMap.StopRun
    End If
    On Error Resume Next
    Maps(activemap).picPreview.setfocus
End Sub






Private Sub cTileset_SelectionChange()
    If loadedmaps(activemap) Then
        Call UpdatePreview
    End If
End Sub

Private Sub customShapeSize_Change(Index As Integer)
    Call SetSetting("CustomShapeSize" & CustomShapeName(Index + 0), CStr(customShapeSize(Index).value))
    Call UpdateCustomShapePreview(Index + 0)
End Sub

Private Sub customShapeTeethNumber_Change(Index As Integer)
    Call SetSetting("CustomShapeTeethNumber" & CustomShapeName(Index + 0), CStr(customShapeTeethNumber(Index).value))
    Call UpdateCustomShapePreview(Index + 0)

End Sub

Private Sub customShapeTeethSize_Change(Index As Integer)
    Call SetSetting("CustomShapeTeethSize" & CustomShapeName(Index + 0), CStr(customShapeTeethSize(Index).value))
    Call UpdateCustomShapePreview(Index + 0)

End Sub

Sub UpdateCustomShapePreview(Index As customshapeEnum)
    If loadedmaps(activemap) Then
        Select Case Index
        Case s_cogwheel
            Call Maps(activemap).tline.DrawCogWheel(0, 0, 0, 0, frmGeneral.customShapeTeethNumber(Index).value, frmGeneral.customShapeTeethSize(Index).value / 100, Nothing, False, False, True)
        Case s_star
            Call Maps(activemap).tline.DrawStar(0, 0, 0, 0, frmGeneral.customShapeTeethNumber(Index).value, frmGeneral.customShapeTeethSize(Index).value / 100, Nothing, False, False, True)
        Case s_regular
            Call Maps(activemap).tline.DrawRegularShape(0, 0, 0, 0, frmGeneral.customShapeTeethNumber(Index).value, Nothing, False, False, True)
        End Select
    End If
End Sub
' ---------------------------------------------------------------
' ---------------------------------------------------------------
' ---------------------- CONTROL EVENTS -------------------------
' ---------------------------------------------------------------
' ---------------------------------------------------------------






Private Sub Label6_Click()
    MessageBox Maps(activemap).lvz.ListSelectedTrue
End Sub

Private Sub lblswitchtiles_Click()
' Switches the right and left tiles with each other
    If loadedmaps(activemap) Then
        Call Maps(activemap).tileset.SwapSelections
    End If
End Sub


Private Sub UpdateRegionList()
    Dim i As Integer
    Dim oldIndex As Integer
    oldIndex = llRegionList.ListIndex

    llRegionList.Clear
    
    If loadedmaps(activemap) Then
        For i = 0 To Maps(activemap).Regions.getRegionIndex
            Call llRegionList.addItem(Maps(activemap).Regions.getRegionName(i), Maps(activemap).Regions.getRegionIsVisible(i), Maps(activemap).Regions.getRegionColor(i))
        Next
    
        If Maps(activemap).Regions.getRegionIndex <> -1 Then
    
            If oldIndex >= 0 And oldIndex <= Maps(activemap).Regions.getRegionIndex Then
                llRegionList.ListIndex = oldIndex
            Else
                llRegionList.ListIndex = 0
            End If
    
        End If
    End If
End Sub

Private Sub llRegionList_AddItemClick()
    If Not loadedmaps(activemap) Then Exit Sub
    
    Dim name As String
    Dim defaultname As String

    Dim i As Integer
    i = 1
    While Maps(activemap).Regions.regionNameExists("New Region " & CStr(i))
        i = i + 1
    Wend
    defaultname = "New Region " & CStr(i)

    name = InputBox("Please enter a name for your new region", "New Region", defaultname)
    If name <> "" Then
        Call Maps(activemap).Regions.NewRegion(name)
        Call UpdateRegionList
        llRegionList.ListIndex = Maps(activemap).Regions.getRegionIndex
    End If
End Sub

Private Sub llRegionList_change()
    If Not loadedmaps(activemap) Then Exit Sub
    
    If llRegionList.ListIndex <> -1 Then
        Call Maps(activemap).Regions.SelectRegion(llRegionList.ListIndex, Not dontRefreshRegions)
        
        Call llRegionList.Redraw
    End If
End Sub

Private Sub llRegionList_ChangeItemColor(Index As Integer)
    If Not loadedmaps(activemap) Then Exit Sub
    
    Call Maps(activemap).Regions.SetColor(Index, GetColor(Me, Maps(activemap).Regions.getRegionColor(Index), True, False))
'20        Call Maps(activemap).Regions.BuildRegionTiles
    dontRefreshRegions = True
    
    Call UpdateRegionList
    llRegionList.ListIndex = Index    'select that region after
    
    dontRefreshRegions = False
    
    Call Maps(activemap).Regions.RedrawAllRegions
    Call Maps(activemap).RedrawRegions(True)
End Sub

Private Sub llRegionList_DeleteItemClick(Index As Integer)
    If Not loadedmaps(activemap) Then Exit Sub
    
    If MessageBox("Delete " & Maps(activemap).Regions.getRegionName(Index) & " ?", vbQuestion + vbYesNo, "Delete Region") = vbYes Then
        Call Maps(activemap).Regions.DeleteRegion(Index)
        Call UpdateRegionList
    End If
End Sub

Private Sub llRegionList_EditItemClick(Index As Integer)
    If Not loadedmaps(activemap) Then Exit Sub
    
    Load frmEditRegion
    Call frmEditRegion.setParent(Maps(activemap), llRegionList.ListIndex)
    Call frmEditRegion.UpdatePreview
    frmEditRegion.show vbModal, Me
End Sub

Private Sub llRegionList_RightClick(Index As Integer)
    Dim oldname As String, newname As String
    
    If loadedmaps(activemap) Then
        oldname = Maps(activemap).Regions.getRegionName(Index)
        Do
            newname = InputBox("Please enter a new name for region", "Rename region '" & Maps(activemap).Regions.getRegionName(Index) & "'", oldname)
        Loop While newname = "" And oldname = ""
    
        If newname <> "" Then
            Call Maps(activemap).Regions.setRegionName(Index, newname)
            Call UpdateRegionList
'            llRegionList.ListIndex = Index
        End If
    End If
    
End Sub

Private Sub llRegionList_VisibiltyChanged(Index As Integer)
    If loadedmaps(activemap) Then
        Call Maps(activemap).Regions.ToggleVisible(Index, True)
'        llRegionList.ListIndex = Index    'select that region after
        Call Maps(activemap).RedrawRegions(True)
    End If
End Sub

Private Sub MDIForm_Activate()
'All to be done is in resize
    Call MDIForm_Resize
    If loadedmaps(activemap) Then
'30            DoEvents
        Call Maps(activemap).UpdateLevel(False, True)
    End If
End Sub

Private Sub MDIForm_Deactivate()
    Call llRegionList.HideList
End Sub


Private Sub ShowDebugInformation(ByVal show As Boolean)

    Label6.visible = show
    Label7.visible = show
    mnuDebug.visible = show
'    mnuSpeedTest.visible = show
'    mnuShowMemUsage.visible = show
    mnuShowDebugInfo.checked = show
End Sub

Private Sub MDIForm_Load()
' Loading of the general form
'set debug to true to avoid complications with subclassing in the IDE
'this statement will only be executed in the IDE
    
    
    Dim i As Integer
    
    If Not bDEBUG Then On Error GoTo MDIForm_Load_Error
    
    
    Call ShowDebugInformation(bDEBUG)
    
    For i = DL_Regions To DL_Buffer
        Set MapLayers(i) = New clsDisplayLayer
    Next
    
'100       Hook Me.hWnd
    
    HookWnd Me.hWnd
    
    If Me.windowstate <> vbMinimized Then
        lastwindowstate = Me.windowstate
    Else
        lastwindowstate = vbMaximized
    End If
    
    'For some reason, there is always a tab by default, let's get rid of it... right here, right now.
    tbMaps.Tabs.Remove "Dummy"
    
    
    'check if we have an instance of ourselves running
'170       EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    
    Me.Caption = "Drake Continuum Map Editor (v" & App.Major & "." & App.Minor & "." & App.Revision & ")"
      
    AddDebug vbNewLine & Now & " --- " & Me.Caption & " @ " & GetApplicationFullPath & " starting...", True
    AddDebug "Windows version: " & currentWindowsVersion
    AddDebug "Screen resolution: " & Screen.width \ Screen.TwipsPerPixelX & " x " & Screen.height \ Screen.TwipsPerPixelY
    
'200       AddDebug Me.hWnd & " +++ OtherInstance " & otherinstance, True
             
    frmGeneral.IsBusy("frmGeneral.MDIForm_Load") = True
    

    Me.show
    SetIconsToMenus

    
    
    
    Dim toolopt As frame
    For Each toolopt In frmTool
        c_toolOptionExists(toolopt.Index + 1) = True
    Next
    
    'load settings from which toolbars are shown
    Toolbarleft.visible = CBool(GetSetting("ShowToolbarTools", "1"))
    toolbartop.visible = CBool(GetSetting("ShowToolbarStandard", "1"))
    tlbTabs.visible = CBool(GetSetting("ShowToolbarMapTabs", "1"))

    'set main tool on hand
    Call SetCustomShapeMenu(CInt(GetSetting("LastUsedCustomShape", "0")))
    SetCurrentTool (T_hand)


    'load the recent entry's
    Call LoadRecent

    'init clipboard
    ReDim SharedVar.clipdata(1024, 1024)
'620       ReDim SharedVar.clipBitField(1024, 1024)



      
    'load the args if we need to load it
'630       If otherinstance = 0 Then

  Dim args As String
  Dim tmp() As String
  args = Command()
  
  AddDebug "args:'" & args & "'"
  
  Dim ret As Integer
  Dim openedmapbyargs As Boolean
  Dim longfilename As String
  Dim fileext As String
  
        tmp = Split(args, Chr(34))
        For i = 0 To UBound(tmp)
            
            AddDebug "Reading '" & tmp(i) & "'"
            
          longfilename = GetLongFilename(tmp(i))
          
          AddDebug "Attempting to loaddd '" & longfilename & "'"
          
          fileext = GetExtension(GetFileTitle(longfilename))
          
            If fileext = "lvl" Or fileext = "elvl" Or fileext = "bak" Then
          'if it's a lvl file, try loading it
                
                AddDebug "Attempting to load '" & longfilename & "'"
                
                openedmapbyargs = True
                ret = OpenMap(GetLongFilename(tmp(i)))
                If ret = -1 Then Exit For

            End If

        Next
'720       End If

    'open a new map automaticly, if there isn't a map loaded by the
    'arguments given
    If (Not openedmapbyargs) Then
        If GetSetting("AutoNewMap", "1") = "1" Then
            Call NewMap
        Else
            Call UpdateToolBarButtons
        End If
    End If

    Dim backuppath As String
    backuppath = GetSetting("DeleteBackup", "")
    If backuppath <> "" And LCase(backuppath) <> LCase(GetApplicationFullPath) Then
        Call SetSetting("DeleteBackup", "")
        DeleteFile (backuppath)
    End If
        
    frmGeneral.IsBusy("frmGeneral.MDIForm_Load") = False
    

    
    'auto update
    If GetSetting("AutoUpdate", 1) = "1" Then
        Dim updateperiod As Integer
        Dim lastupdate As Date

        quickupdate = True

        lastupdate = CDate(Format(GetSetting("AutoUpdateLast", -1), "dd-mm-yyyy"))
        updateperiod = CInt(val(GetSetting("AutoUpdateDelay", 1)))

        AddDebug "+++ Last update: " & Format(lastupdate, "dd-mm-yyyy") & " - Update period: " & updateperiod

        If lastupdate <> -1 And Not updateperiod = 0 Then

            'compare last update to current date
            If (Date - lastupdate >= 1 And updateperiod = 1) Or _
               (Date - lastupdate >= 7 And updateperiod = 2) Or _
               (Date - lastupdate >= 30 And updateperiod = 3) Then
                'update
                Call LoadUpdateForm
            End If
        ElseIf lastupdate = -1 Then
            Call SetSetting("AutoUpdateLast", "0")
            'it was never updated yet
            'so: update immediatly
'610               If inSplash Then
'620                 MakeNormal frmSplash.hWnd
'630               End If
            
            If MessageBox("DCME will now automatically check for the newest version available.", vbOKCancel, "Update") = vbYes Then
                Call LoadUpdateForm
            End If

'670               If inSplash Then
'680                 MakeTopMost frmSplash.hWnd
'690               End If
        Else
            'It is set to check for updates everytime it starts
            Call LoadUpdateForm

        End If

        quickupdate = False
    Else
        AddDebug "+++ Skipped check for updates"
    End If

    MAX_TOOL_SIZE(T_pencil - 1) = 64
    MAX_TOOL_SIZE(T_line - 1) = 64
    MAX_TOOL_SIZE(t_spline - 1) = 64
    MAX_TOOL_SIZE(T_rectangle - 1) = 64
    MAX_TOOL_SIZE(T_ellipse - 1) = 64
    MAX_TOOL_SIZE(T_filledellipse - 1) = 64
    MAX_TOOL_SIZE(T_filledrectangle - 1) = 64
    toolSize(T_pencil - 1).Max = 64
    toolSize(T_line - 1).Max = 64
    toolSize(t_spline - 1).Max = 64
    toolSize(T_rectangle - 1).Max = 64
    toolSize(T_ellipse - 1).Max = 64
    toolSize(T_filledellipse - 1).Max = 64
    toolSize(T_filledrectangle - 1).Max = 64

    toolStep(T_line - 1).Min = 0
    toolStep(t_spline - 1).Min = 0
    toolStep(T_rectangle - 1).Min = 0
    toolStep(T_ellipse - 1).Min = 0
    toolStep(T_filledellipse - 1).Min = 0
    toolStep(T_filledrectangle - 1).Min = 0


    customShapeSize(s_cogwheel).Min = 1
    customShapeTeethNumber(s_cogwheel).Min = 3
    customShapeTeethSize(s_cogwheel).Min = 0
    customShapeSize(s_cogwheel).Max = 64
    customShapeTeethNumber(s_cogwheel).Max = 50
    customShapeTeethSize(s_cogwheel).Max = 100


    customShapeSize(s_star).Min = 1
    customShapeTeethNumber(s_star).Min = 3
    customShapeTeethSize(s_star).Min = 0
    customShapeSize(s_star).Max = 64
    customShapeTeethNumber(s_star).Max = 50
    customShapeTeethSize(s_star).Max = 100


    customShapeSize(s_regular).Min = 1
    customShapeTeethNumber(s_regular).Min = 3
    customShapeSize(s_regular).Max = 64
    customShapeTeethNumber(s_regular).Max = 50

    sldAirbSize.Max = MAX_AIRBR_SIZE
    sldAirbDensity.Max = MAX_AIRBR_DENSITY
    
    Dim tp As ToolProperty
    For Each tp In toolSize
        tp.Caption = "Size"
    Next
    For Each tp In toolStep
        tp.Caption = "Step"
    Next

    For Each tp In customShapeSize
        tp.Caption = "Size"
    Next
    For Each tp In customShapeTeethNumber
        tp.Caption = "Teeth nr"
    Next
    For Each tp In customShapeTeethSize
        tp.Caption = "Teeth size"
    Next
    Call UpdateToolToolbar
    
    
    Call cTileset.Redraw
    
    Me.visible = True
    Unload frmSplash
      
      'show tips
    If GetSetting("ShowTips", 1) = "1" Then
        Call mnuTips_Click
    End If

    AddDebug Now & " --- DCME ready"
    
    On Error GoTo 0
    Exit Sub
MDIForm_Load_Error:
    HandleError Err, "MDIForm_Load", True, True
End Sub


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If AutoHideTileset And FloatTileset.visible = True Then
        HideFloatTileset
    End If
    If AutoHideRadar And FloatRadar.visible = True Then
        HideFloatRadar
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Try to close every map
'Release the subclassing hook
    On Error GoTo MDIForm_QueryUnload_Error
    
'    If PRINTDEBUG Then
'        AddDebug "At MDIForm_QueryUnload " & Cancel & " " & UnloadMode
'    End If
    If inSplash Then Unload frmSplash
    
    Dim i As Integer
    For i = 0 To UBound(Maps)
        If loadedmaps(i) Then
            'Maps(i).SetFocus
            Dim ret As Integer
            ret = False

            Call Maps(i).Form_QueryUnload(ret, QueryUnloadConstants.vbFormControlMenu)

            If ret = True Then
                Cancel = True
                Call UpdateMenuMaps
                Exit Sub
            Else
                Call DestroyMap(i)
            End If
        End If
    Next
    
    If updateready = True Then

        'check if the zip file is still present, abort if not
        If FileExists(updatefilepath) Then

            On Error GoTo Update_Error
            
            'an update was made, prepare the file for the next startup
            'clear the caption so the new exe does not find it
            frmGeneral.Caption = "Updating..."

            AddDebug "+++ Updating with " & updatefilepath

            'rename current exe
            Dim sSourceFile As String
            Dim sDestinationFile As String

            sSourceFile = GetApplicationFullPath
            sDestinationFile = App.path & "\DCMEbackup.exe"
            
            AddDebug "+++ Renaming " & sSourceFile & " to " & sDestinationFile
            
            Call RenameFile(sSourceFile, sDestinationFile)
                        
            'take note to delete that backup copy on next startup
            Call SetSetting("DeleteBackup", sDestinationFile)
            
            AddDebug "+++ Starting " & updatefilepath

            'check again if the file exists.. just to be sure
            If FileExists(updatefilepath) Then
                'execute the self-extracting update archive and overwrite

                ShellExecute 0&, vbNullString, updatefilepath, vbNullString, GetPathTo(updatefilepath), vbHide
                'the extracted exe will be started automatically
                'by the archive itself
            Else
              AddDebug "+++ " & updatefilepath & " does not exist"
            End If
            'continue unloading old exe
            
            On Error GoTo MDIForm_QueryUnload_Error
        End If

    End If


MDIForm_QueryUnload_End:
    'save settings before ending (for toolbar status etc
    Call SaveSettings
    
    'make sure floating windows are unloaded
    'They shouldn't be a problem anymore, but there's still a
    'small risk of like... hitting alt-f4 while they're popped up or something
    'so let's unload them anyway
    Call HideFloatRadar
    Call HideFloatTileset
    Unload frmTip
'410       Unload dlgProgress
    Unload frmTilesetEditor
    
    'TODO: These shouldn't be unloaded here
    'make sure createwalltiles is unloaded
    Unload frmCreateWallTile
    
    Unload frmSave
    
    'Clear cache
    Call DeleteDirectory(Directory_Cache)
    
    
    
    'check if any form is remaining

    For i = DL_Regions To DL_Buffer
        Set MapLayers(i) = Nothing
    Next
    
    If CheckUnloadedForms(True) Then
        AddDebug "DCME ended successfully"
    End If
  
    
  
  
    UnHookWnd Me.hWnd
'410       UnHook Me.hWnd
    Unload Me

    Exit Sub
MDIForm_QueryUnload_Error:
    'critical error (no need to unload me, handleerror will do that)
    HandleError Err, "frmGeneral.QueryUnload " & Cancel & " " & UnloadMode, True, True
    
    Exit Sub
Update_Error:
    If Err.Number = 429 Then
        'ActiveX component can't create object.
        MessageBox "DCME could not complete the update. Please execute " & GetFileTitle(updatefilepath) & " once DCME is terminated.", vbOKOnly + vbExclamation, "Manual update required"
    Else
        HandleError Err, "frmGeneral.QueryUnload (update)", True, False
    End If
    GoTo MDIForm_QueryUnload_End
    
End Sub

Function CheckUnloadedForms(showMsgBox As Boolean) As Boolean
'Returns true if all forms (except frmGeneral) are unloaded already
'Unloads all these forms
    CheckUnloadedForms = True
    
    Dim names As String
    Dim count As Integer
    
    
    Dim f As Form
    For Each f In Forms
        If Not (f Is Me) Then
            count = count + 1
            If count > 1 Then names = names & ", "
            names = names & f.name
            AddDebug f.name & " was not unloaded correctly."
            
            CheckUnloadedForms = False
            Unload f
        End If
    Next
    
    If showMsgBox And count > 0 Then
        If count = 1 Then
            MessageBox names & " was not unloaded correctly.", vbCritical + vbOKOnly
        ElseIf count > 1 Then
            MessageBox count & " items were not unloaded correctly: " & names, vbCritical + vbOKOnly
        End If
    End If
End Function


Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LoadMapFromOLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub MDIForm_Resize()
'Resize picradar when resized
'calculate dimensions picrightbar
'don't resize the radar when hiding the right bar, this will screw up the popup radar

    On Error GoTo MDIForm_Resize_Error
'    If PRINTDEBUG Then
'        AddDebug "At MDIForm_Resize (windowstate=" & Me.windowstate & ")"
'    End If
    
    

    
    Dim i As Integer
    
    If Me.windowstate = 0 And Me.Top < 0 Then Me.Top = 0

    If Not dontResizeRadar Then
        Dim totalWidth As Integer
        totalWidth = frmGeneral.picRightBar.width / Screen.TwipsPerPixelX
        Dim totalHeight As Integer
        totalHeight = frmGeneral.picRightBar.height / Screen.TwipsPerPixelY

        'set picradar width = width of right bar, maintain square
        picradar.width = totalWidth
        picradar.height = picradar.width

        'if the height is larger than the amount of space available
        'resize until it fits (maintain square)
        If (picradar.Top + picradar.height) > totalHeight - 17 Then
            Dim h As Integer
            h = picradar.height - ((picradar.Top + picradar.height) - (totalHeight - 17))
            If h < 100 Then
                picradar.height = 100
            Else
                picradar.height = picradar.height - ((picradar.Top + picradar.height) - (totalHeight - 17))
            End If
            picradar.width = picradar.height
        End If
        
        chkForceFullRadar.Top = picradar.Top + picradar.height

        'now center the radar
        picradar.Left = totalWidth / 2 - picradar.width / 2
    End If

    If loadedmaps(activemap) Then
        Call Maps(activemap).UpdateScrollbars(False)
    End If

    'make the toolbar for the tools = the length of the toolbar
    Dim frm As frame
    For Each frm In frmTool
        frm.width = tlbToolOptions.width
        frm.Left = 0
        frm.Top = 0
    Next
    
'    For i = frmTool.LBound To frmTool.UBound
'        If toolOptionExist(i + 1) Then
'            frmTool(i).Width = tlbToolOptions.Width
'            frmTool(i).Left = 0
'            frmTool(i).Top = 0
'        End If
'    Next
    
    For Each frm In frmCustomShape
        frm.width = tlbToolOptions.width
        frm.Left = 0
        frm.Top = 0
    Next
    
'    For i = frmCustomShape.LBound To frmCustomShape.UBound
'        If customShapeOptionExist(i + 0) Then
'            frmCustomShape(i).Width = tlbToolOptions.Width
'            frmCustomShape(i).Left = 0
'            frmCustomShape(i).Top = 0
'        End If
'    Next
    

    
    If Me.windowstate <> vbMinimized Then
        If tlbTabs.width - tbMaps.Left - 120 > 10 Then
            tbMaps.width = tlbTabs.width - tbMaps.Left - 120
        End If
        
        If loadedmaps(activemap) Then
          Call Maps(activemap).Form_Resize_Force
          Call Maps(activemap).UpdateLevel(False, True)
        End If
'370           For i = 0 To 9
'380               If loadedmaps(i) Then
'390                   Call Maps(i).Form_Resize_Force
'400                   Call Maps(i).UpdateLevel(False, True)
'410               End If
'420           Next
    
        lastwindowstate = Me.windowstate
    End If
    
    Call UpdateToolToolbar
    Call llRegionList.HideList
    
    On Error GoTo 0
    Exit Sub
    
MDIForm_Resize_Error:
    HandleError Err, "MDIForm_Resize"
End Sub





Private Sub mnuAbout_Click()
'Shows the About Menu
    frmSplash.show vbModal, frmGeneral
End Sub

Private Sub mnuAddAll_Click()
'Add all tiles in screen to selection
    If Not loadedmaps(activemap) Then Exit Sub

    Call Maps(activemap).sel.SelectAllTiles(True, True)
End Sub

Private Sub mnuAddLeftSelectedTile_Click()
'Add all left tiles in screen to the selection
    If Not loadedmaps(activemap) Then Exit Sub
    
    If Maps(activemap).tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.AddTiles(Maps(activemap).tileset.selection(vbLeftButton).tilenr, True)
    End If
End Sub

Private Sub mnuAddLeftSelectedTileMap_Click()
'Add all left tiles on the map to the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.AddTiles(Maps(activemap).tileset.selection(vbLeftButton).tilenr, False)
    End If
End Sub

Private Sub mnuAddRightSelectedTile_Click()
'Add all right tiles in screen to the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.AddTiles(Maps(activemap).tileset.selection(vbRightButton).tilenr, True)
    End If
End Sub

Private Sub mnuAddRightSelectedTileMap_Click()
'Add all right tiles on the map to the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.AddTiles(Maps(activemap).tileset.selection(vbRightButton).tilenr, False)
    End If
End Sub



Private Sub mnuBookmark_Click(Index As Integer)
    Call Maps(activemap).GotoBookMark(Index)
End Sub

Private Sub mnuCascade_Click()
'Cascades the windows
    Me.Arrange vbCascade
End Sub

Private Sub mnuCenterSelection_Click()
'Center the selection in the screen
    If Not loadedmaps(activemap) Then Exit Sub

    Call Maps(activemap).sel.DoCenterSelection(True)

End Sub

Private Sub mnuCenterSelectionMap_Click()
'Center the selection on the map
    If Not loadedmaps(activemap) Then Exit Sub


    Call Maps(activemap).sel.DoCenterSelection(False)

End Sub

Private Sub mnuCheckUpdates_Click()
'Shows the check for updates form
    If Not updateformloaded Then
        Call LoadUpdateForm(True)
    Else
        frmCheckUpdate.visible = True
        frmCheckUpdate.setfocus
        quickupdate = False
    End If
End Sub

Private Sub mnuClose_Click()
'Close the activemap
    If Not loadedmaps(activemap) Then Exit Sub

    Call DestroyMap(activemap)
End Sub

Private Sub mnuConvtowalltiles_Click()
'Show the Convert to walltiles form
    If Not loadedmaps(activemap) Then Exit Sub

    Call frmConvToWallTile.setParent(Maps(activemap))
    frmConvToWallTile.show vbModal, frmGeneral
End Sub

Private Sub mnuCopy_Click()
'Copies the selection if a selection is active
    If Not loadedmaps(activemap) Then Exit Sub

    Call Maps(activemap).sel.CopySelection
End Sub

Private Sub mnuCount_Click()
'Counts the tiles
    Call ExecuteCount
End Sub

Sub SetCustomShapeMenu(Index As Integer)
    curCustomShape = Index
    Dim i As Integer
    For i = mnuCustomShape.LBound To mnuCustomShape.UBound
        mnuCustomShape(i).checked = False
    Next

    mnuCustomShape(Index).checked = True
End Sub
Private Sub mnuCustomShape_Click(Index As Integer)
    Call SetCustomShapeMenu(Index)

    Call SetSetting("LastUsedCustomShape", CStr(Index))

    Call setToolToolbar(T_customshape)
    Call UpdateToolToolbar
    
    Call UpdateCustomShapePreview(Index + 0)

End Sub

Private Sub mnuCut_Click()
'Cuts the selection if a selection is active
    If Not loadedmaps(activemap) Then Exit Sub
    'call the cut selection of the map's selection
    Call Maps(activemap).sel.CutSelection
End Sub



Private Sub mnuDebugLayers_Click()
    If Not loadedmaps(activemap) Then Exit Sub
    
    Call Maps(activemap).ExportLayers
    
    
End Sub

Private Sub mnuDebugLog_Click()
    If FileExists(App.path & "\DCME.log") Then
'        Shell "Notepad " & App.path & "\DCME.log", vbNormalFocus
        ShellExecute 0&, vbNullString, "notepad", App.path & "\DCME.log", vbNullString, vbNormalFocus
    End If
End Sub

Private Sub mnuDeleteSelection_Click()
'Deletes the selection if there is a selection
    If Not loadedmaps(activemap) Then Exit Sub
    If Not Maps(activemap).sel.hasAlreadySelectedParts Then Exit Sub

    Call Maps(activemap).ClearSelection
End Sub

Private Sub mnudiscardtileset_Click()
'Discards the current tileset
    If Not loadedmaps(activemap) Then Exit Sub

    'discard the tileset of the active map
    Call Maps(activemap).DiscardTileset
End Sub

Private Sub mnuEditTileset_Click()
'Shows the Tileset Editor
    If Not loadedmaps(activemap) Then Exit Sub

    Load frmTilesetEditor
    frmTilesetEditor.tilesetpath = Maps(activemap).tilesetpath
    frmTilesetEditor.lblTileset = "Tileset - ''" & Maps(activemap).tilesetpath & "''"
    Call frmTilesetEditor.InitLeftSelection(Maps(activemap).tileset.selection(vbLeftButton))
    frmTilesetEditor.visible = True
End Sub

Private Sub mnuElvl_Click()
    If Not loadedmaps(activemap) Then Exit Sub

    Call frmeLVL.setParent(Maps(activemap))
    frmeLVL.show vbModal, frmGeneral

    Call UpdateToolToolbar
End Sub

Private Sub mnuExit_Click()
'Exits the program
    Unload Me
End Sub

Private Sub mnuExportTileset_Click()
'Exports the tileset
    Call ExportTileset
End Sub

Private Sub mnuFlip_Click()
'Flips the selection if a selection is active
    If Not loadedmaps(activemap) Then Exit Sub
    'flip the selection
    Call ExecuteFlip
End Sub

Private Sub mnuGrid_Click()
'Toggles the grid
    Call ToggleGrid
End Sub

Private Sub mnuImportTileset_Click()
'Imports a Tileset
    Call ImportTileset
End Sub

Private Sub mnulstAutosaves_Click(Index As Integer)
    Call OpenMap(mnulstAutosaves(Index).Tag)
    
End Sub

Private Sub mnulstRecent_Click(Index As Integer)
'Open a recent file
'recent menus are "" when they contain no data,
'and we cannot hide mnu 0 so we need to check if
'its <> ""
    If mnulstRecent(Index).Caption <> "" Then
        'check if the file still exists
        If FileExists(mnulstRecent(Index).Caption) Then
            'open the map, entry in openmap for recent menus wont'
            'be executed because it is already in the recent list
            Dim ret As Integer
            ret = OpenMap(mnulstRecent(Index).Caption)
            If ret < 0 Then
                Exit Sub
            End If

            'move the recent to the top
            Call MoveRecentToTop(Index)
        Else
            'doesnt exist anymore
            MessageBox "Error: The file does not longer exist!", vbExclamation
        End If
    End If
End Sub

Private Sub mnuLvzAddImage_Click()
    If Not loadedmaps(activemap) Then Exit Sub
    
    Call frmLVZAddImage.setParent(Maps(activemap))
    frmLVZAddImage.show vbModal, frmGeneral
End Sub

Private Sub mnuLvzDeleteImage_Click()
    If Not loadedmaps(activemap) Then Exit Sub
    
    With Maps(activemap)
    
        If .tileset.selection(1).selectionType = TS_LVZ Then
            'We have an image selected
        
            Dim lvzidx As Integer, imgidx As Integer, globalId As Integer
        
            lvzidx = .tileset.selection(1).group
            imgidx = .tileset.selection(1).tilenr
        
            globalId = .lvz.GetGlobalIndexOfImageDefinition(lvzidx, imgidx)
            
            
            Dim MapObjCount As Long, ScrObjCount As Long
            
            MapObjCount = .lvz.CountMapObjectsUsingImage(lvzidx, imgidx)
            ScrObjCount = .lvz.CountScreenObjectsUsingImage(lvzidx, imgidx)
            
            If MapObjCount > 0 Or ScrObjCount > 0 Then
                Dim msgret As VbMsgBoxResult
                msgret = MessageBox("There are " & MapObjCount & " map objects and " & ScrObjCount & " screen objects associated to this image. Deleting it will also delete all these objects. Do you wish to delete it?", vbExclamation + vbYesNo, "Confirm Delete Image Definition")
                
                If msgret = vbNo Then
                    Exit Sub
                End If
                
            End If
            
            'Do it
            Call .lvz.RemoveLinksToImage(lvzidx, imgidx)
            Call .lvz.removeImageDefinitionFromLVZ(lvzidx, imgidx)
            
            
            'Find the next image in the library
            Call .lvz.GetLocalIndexOfImageDefinition(globalId, lvzidx, imgidx)
            
            If lvzidx >= 0 And imgidx >= 0 Then
                Call .tileset.SelectLVZ(vbLeftButton, lvzidx, imgidx, False)
            Else
                Call .tileset.SelectTiles(vbLeftButton, 1, 1, 1, False)
            End If

            
'            Call picLVZImages_MouseDown(vbLeftButton, 0, CSng(clickX), CSng(clickY))
            
            'Some objects might have been deleted
            Call .UpdateLevel
            
            Call cTileset.DrawLVZTileset(True)
'            Call .tileset.DrawLVZTileset(True)
                
            
        End If
    
    End With
    
End Sub

Private Sub mnuLvzEditAnimation_Click()
    If Not loadedmaps(activemap) Then Exit Sub
    
    Dim lvzidx As Integer, imgidx As Integer

    
    With Maps(activemap)

        lvzidx = .tileset.selection(1).group
        imgidx = .tileset.selection(1).tilenr
        
        If lvzidx >= 0 And imgidx >= 0 Then
            Call .lvz.DoEditImageDefinitionProperties(lvzidx, imgidx)
        End If
        
        Call .UpdateLevel
        
        
'        Call .tileset.DrawLVZTileset(True)
    
    End With
    
    Call cTileset.DrawLVZTileset(True)
    
End Sub

Private Sub mnuLvzEditImage_Click()
    If Not loadedmaps(activemap) Then Exit Sub

    Dim lvzidx As Integer, imgidx As Integer

    
    With Maps(activemap)
    
        lvzidx = .tileset.selection(1).group
        imgidx = .tileset.selection(1).tilenr
        
        If lvzidx >= 0 And imgidx >= 0 Then
            Call .lvz.DoEditImageDefinitionPic(lvzidx, imgidx)
        End If
        
        Call .UpdateLevel
        
    End With
    
    Call cTileset.DrawLVZTileset(True)
'        Call .tileset.DrawLVZTileset(True)
        
    
    
End Sub

Private Sub mnuLvzOpenManager_Click()
    Call mnuManageLVZ_Click
End Sub


Private Sub mnuManageLVZ_Click()
    If Not loadedmaps(activemap) Then Exit Sub
    
    Call frmLVZ.setParent(Maps(activemap))
    frmLVZ.show vbModal, frmGeneral
    'frmLVZ.Show
End Sub

Private Sub mnuMaps_Click(Index As Integer)
      'Shows the selected map
      'show the current selected map in the map menus under window
    Call ActivateMap(Index)
          
End Sub

Sub ActivateMap(Index As Integer)
    If activemap <> Index And loadedmaps(Index) Then
        Maps(activemap).Form_Deactivate
        Maps(Index).setfocus
        'stupid thing doesn't call form_activate when maximized
        Maps(Index).Form_Activate
    End If
End Sub


Private Sub mnuMirror_Click()
'Mirrors the selection
    If Not loadedmaps(activemap) Then Exit Sub

    Call ExecuteMirror
End Sub

Private Sub mnuNew_Click()
'Creates a new map
    Call NewMap
End Sub

Private Sub mnuOpen_Click()
'Opens a map, no path is given so a common dialog should appear
    Call OpenMap("")
End Sub

Private Sub mnuPaste_Click()
'Pastes the data from the clipboard
    If Not loadedmaps(activemap) Then Exit Sub
    'paste into a new selection
    Call frmGeneral.clipboard.Paste(Maps(activemap).sel)
End Sub

Private Sub mnuPreferences_Click()
'Show the options form
          frmOptions.show vbModal, Me
End Sub

Private Sub mnuPTM_Click()
'Shows the Picture To Map form, and passes the tilesetleft to that form
    If Not loadedmaps(activemap) Then Exit Sub

'    If Maps(activemap).tileset.selection(vbLeftButton).is.tilesetleft = 217 Or Maps(activemap).tilesetleft = 219 Or Maps(activemap).tilesetleft = 220 Then
'        'cant use tiles bigger than 1 tile in PTM
'        messagebox "You can't use this tool with special tiles.", vbInformation
'        Exit Sub
'    Else
    'shows the form and sets the tiletouse to the left tile
    Load frmPicToMap
    Call frmPicToMap.setParent(Maps(activemap))
    
    frmPicToMap.show vbModal, frmGeneral

End Sub

Private Sub mnuRedo_Click()
'Redo the undo if available
    If Not loadedmaps(activemap) Then Exit Sub
    'do redo
    Call Maps(activemap).undoredo.Redo
End Sub

Private Sub mnuRemoveLeftSelectedTileMap_Click()
'Removes the left selected tile on the map from the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.RemoveTiles(Maps(activemap).tileset.selection(vbLeftButton).tilenr, False)
    End If
End Sub

Private Sub mnuRemoveRightSelectedTile_Click()
'Removes the right selected tile on the map from the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.RemoveTiles(Maps(activemap).tileset.selection(vbRightButton).tilenr, True)
    End If
End Sub

Private Sub mnuRemoveLeftSelectedTile_Click()
'Removes the left selected tile in the screen from the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.RemoveTiles(Maps(activemap).tileset.selection(vbLeftButton).tilenr, True)
    End If
End Sub

Private Sub mnuRemoveRightSelectedTileMap_Click()
'Removes the left selected tile in the screen from the selection
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.RemoveTiles(Maps(activemap).tileset.selection(vbRightButton).tilenr, False)
    End If
End Sub

Private Sub mnuReplace_Click()
'Shows the switch/replace window
    If Not loadedmaps(activemap) Then Exit Sub
    
    Load frmReplace
    
    Call frmReplace.setParent(Maps(activemap))
    
    Call frmReplace.InitSelections
    
    'if we have a selection, auto select the switch/replace in selection
    If Maps(activemap).sel.hasAlreadySelectedParts Then
        frmReplace.chkinselection.Enabled = True
        frmReplace.chkinselection.value = vbChecked
    Else
        frmReplace.chkinselection.Enabled = False
        frmReplace.chkinselection.value = vbUnchecked
    End If

    'show the replace form
    frmReplace.show vbModal, frmGeneral

End Sub

Private Sub mnuResize_Click()
'Shows the resize form
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).sel.hasAlreadySelectedParts Then
        Call frmResize.setParent(Maps(activemap))
        frmResize.show vbModal, frmGeneral
    Else
        'no selection
    End If
End Sub

Private Sub mnuRevert_Click()
    If Not loadedmaps(activemap) Then Exit Sub
    
    If Maps(activemap).activeFile <> "" And Maps(activemap).mapchanged Then
        If MessageBox("Do you wish to revert all changes made to " & Maps(activemap).Caption & "?", vbYesNo + vbQuestion, "Revert all changes") = vbYes Then
            Call LoadRevert
        End If
    End If
    
End Sub

Private Sub mnuRotate_Click()
'Shows the rotate window
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).sel.hasAlreadySelectedParts Then
        'show the rotate form
        frmRotate.show vbModal, frmGeneral
    Else
        'not a valid selection
    End If
End Sub

Private Sub mnuSave_Click()
'Saves the map, no save as
    Call SaveMap(False)
End Sub

Private Sub mnuSaveAs_Click()
'Saves the map, use save as
    Call SaveMap(True)
End Sub

Private Sub mnuSaveMiniMap_Click()
'Saves mini map
    If Not loadedmaps(activemap) Then Exit Sub


    'BitBlt frmSaveRadar.piclevel.hdc, 0, 0, 1024, 1024, Maps(activemap).pic1024.hdc, 0, 0, vbSrcCopy
    Call Maps(activemap).cpic1024.bltToDC(frmSaveRadar.piclevel.hDC, 0, 0, 1024, 1024, 0, 0, vbSrcCopy)
    frmSaveRadar.show vbModal, Me

End Sub



Private Sub mnuSaveSelect_Click()
    If Not loadedmaps(activemap) Then Exit Sub

    frmSave.show vbModal, Me
End Sub

Private Sub mnuSelectAll_Click()
'Select all tiles in the screen
    If Not loadedmaps(activemap) Then Exit Sub

    Call Maps(activemap).sel.SelectAllTiles(True, False)
End Sub

Private Sub mnuSelectAllMap_Click()
'Select all tiles on the map
    If Not loadedmaps(activemap) Then Exit Sub

    Call Maps(activemap).sel.SelectAllTiles(False, False)
End Sub

Private Sub mnuSelectLeftSelectedTile_Click()
'Select left tiles on the screen
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.SelectTiles(Maps(activemap).tileset.selection(vbLeftButton).tilenr, True)
    End If
End Sub

Private Sub mnuSelectLeftSelectedTileMap_Click()
'Select left tiles on the map
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.SelectTiles(Maps(activemap).tileset.selection(vbLeftButton).tilenr, False)
    End If
End Sub

Private Sub mnuSelectNone_Click()
'Discard selection
    If Not loadedmaps(activemap) Then Exit Sub
    Dim undoch As New Changes
    Maps(activemap).undoredo.ResetRedo
    Call Maps(activemap).sel.ApplySelection(undoch, True)

    Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTNONE)
End Sub

Private Sub mnuSelectNoneVisible_Click()
'Deselect all tiles in the screen
    If Not loadedmaps(activemap) Then Exit Sub

    Dim lbx As Integer    'lowestval for x
    Dim hbx As Integer    'highestval for x
    Dim lby As Integer    'lowestval for y
    Dim hby As Integer    'highestval for y
    lbx = Int(Maps(activemap).hScr.value / (Maps(activemap).currenttilew))
    hbx = Int((Maps(activemap).hScr.value + Maps(activemap).picPreview.width) / (Maps(activemap).currenttilew))
    lby = Int(Maps(activemap).vScr.value / (Maps(activemap).currenttilew))
    hby = Int((Maps(activemap).vScr.value + Maps(activemap).picPreview.height) / (Maps(activemap).currenttilew))

    Dim undoch As New Changes
    Maps(activemap).undoredo.ResetRedo

    Call Maps(activemap).sel.RemoveSelectionArea(lbx, hbx, lby, hby, undoch)

    Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTNONE_INSCREEN)
End Sub

Private Sub mnuSelectRightSelectedTile_Click()
'Select right tiles on the screen
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.SelectTiles(Maps(activemap).tileset.selection(vbRightButton).tilenr, True)
    End If
End Sub

Private Sub mnuSelectRightSelectedTileMap_Click()
'Select right tiles on the map
    If Not loadedmaps(activemap) Then Exit Sub

    If Maps(activemap).tileset.selection(vbRightButton).selectionType = TS_Tiles Then
        Call Maps(activemap).sel.SelectTiles(Maps(activemap).tileset.selection(vbRightButton).tilenr, False)
    End If
End Sub

Private Sub mnuSetBookmark_Click(Index As Integer)
    Call Maps(activemap).SetBookMark(Index)
End Sub

Private Sub mnuShowDebugInfo_Click()
    mnuShowDebugInfo.checked = Not mnuShowDebugInfo.checked
    Call ShowDebugInformation(mnuShowDebugInfo.checked)
End Sub

Private Sub mnuShowLVZ_Click()
    Call ToggleLVZ
End Sub

Private Sub mnuShowMemUsage_Click()
    If bDEBUG Then
        MsgBox GetProcessMemory("vb6.exe")
    Else
        MsgBox GetProcessMemory("dcme.exe")
    End If
End Sub

Private Sub mnuShowRegions_Click()
    Call ToggleRegions
End Sub

Private Sub mnuSpeedTest_Click()
    
    Call SpeedTest

End Sub

Private Sub mnuTextToMap_Click()
'Shows the Text To Map form
    If Not loadedmaps(activemap) Then Exit Sub
    
    Load frmTextToMap
    Call frmTextToMap.setParent(Maps(activemap))
    frmTextToMap.show vbModal, frmGeneral
End Sub


Private Sub mnuTileText_Click()
    If Not loadedmaps(activemap) Then Exit Sub

    Load frmTileText
    Call frmTileText.setParent(Maps(activemap))
  Call frmTileText.Init
    frmTileText.show vbModal, Me
End Sub



Private Sub mnuTileH_Click()
'Tiles the windows horizontally
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileNR_Click()
'Toggles the Tile numbers
    ToggleTileNr
End Sub

Private Sub mnuTileV_Click()
'Tiles the windows vertically
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuTips_Click()
'Show tips
    Load frmTip
'    frmTip.forceLoad = True
    frmTip.show vbModal, Me
End Sub

Private Sub mnuTogglePinToolOptions_Click()
    If mnuTogglePinToolOptions.checked Then
        Call SetSetting("PinToolOptions", "0")
    Else
        Call SetSetting("PinToolOptions", "1")
    End If

    Call UpdateToolBarButtons
    Call setToolToolbar(curtool)
End Sub

Private Sub mnuToolbarMapTabs_Click()
    If mnuToolbarMapTabs.checked Then
        Call SetSetting("ShowToolbarMapTabs", "0")
    Else
        Call SetSetting("ShowToolbarMapTabs", "1")
    End If

    Call UpdateToolBarButtons

    'update the scrollbars because the size of the preview changed
    Call Maps(activemap).UpdateScrollbars(True)
End Sub

Private Sub mnuToolbarStandard_Click()
'Toggle visibility of the standard toolbar
    If mnuToolbarStandard.checked Then
        Call SetSetting("ShowToolbarStandard", "0")
    Else
        Call SetSetting("ShowToolbarStandard", "1")
    End If

    Call UpdateToolBarButtons

    'update the scrollbars because the size of the preview changed
    Call Maps(activemap).UpdateScrollbars(True)
End Sub

Private Sub mnuToolbarToolOptions_Click()
'Toggle visibility of the tool options toolbar
    If mnuToolbarToolOptions.checked Then
        Call SetSetting("ShowToolbarToolOptions", "0")
    Else
        Call SetSetting("ShowToolbarToolOptions", "1")
    End If

    Call UpdateToolBarButtons

    'update the scrollbars because the size of the preview changed
    Call Maps(activemap).UpdateScrollbars(True)
End Sub

Private Sub mnuToolbarTools_Click()
'Toggle visibility of the tools toolbar
    If mnuToolbarTools.checked Then
        Call SetSetting("ShowToolbarTools", "0")
    Else
        Call SetSetting("ShowToolbarTools", "1")
    End If

    Call UpdateToolBarButtons

    'update the scrollbars because the size of the preview changed
    Call Maps(activemap).UpdateScrollbars(True)
End Sub

Private Sub mnuTransparentPaste_Click()
'Paste style = Transparent Paste
'update the menu
'    frmGeneral.mnuTransparentPaste.checked = True
'    frmGeneral.mnuPasteUnder.checked = False
'    frmGeneral.mnuNormalPaste.checked = False

    'change the paste type
    Call TogglePasteType(enumPasteType.p_trans)
End Sub

Private Sub mnuNormalPaste_Click()
'Paste style = Normal Paste
'update the menu
'    frmGeneral.mnuTransparentPaste.checked = False
'    frmGeneral.mnuPasteUnder.checked = False
'    frmGeneral.mnuNormalPaste.checked = True

    'change the paste type
    Call TogglePasteType(enumPasteType.p_normal)
End Sub

Private Sub mnuPasteUnder_Click()
'Paste style = Paste Under
'update the menu
'    frmGeneral.mnuTransparentPaste.checked = False
'    frmGeneral.mnuPasteUnder.checked = True
'    frmGeneral.mnuNormalPaste.checked = False

    'change the paste type
    Call TogglePasteType(enumPasteType.p_under)
End Sub

Private Sub mnuUndo_Click()
'Undo the last operation
    If Not loadedmaps(activemap) Then Exit Sub
    'undo any last operations
    Call Maps(activemap).undoredo.Undo
End Sub

Private Sub mnuWalltiles_Click()
' Show the 'Create wall tiles' form
'if no map is active, don't do anything
    Call DoEditWalltiles
End Sub



Sub DoEditWalltiles()
    If loadedmaps(activemap) Then

        Load frmCreateWallTile
        
        Call Maps(activemap).InitfrmWalltiles
        frmCreateWallTile.show vbModal, frmGeneral
    
        Call UpdateToolBarButtons
        
    End If
End Sub

Private Sub optRegionSel_Click(Index As Integer)
    If optRegionSel(REGION_MAGICWAND).value = True Then
        Call SetSetting("RegionUseWand", "1")
        Call SetCurrentCursor(T_magicwand)
    Else
        Call SetSetting("RegionUseWand", "0")
        Call SetCurrentCursor(T_selection)
    End If
End Sub

Private Sub optShip_GotFocus(Index As Integer)
    optShip(Index).value = True
    Call SetSetting("TestMapShip", CStr(Index + 1))

    Call Maps(activemap).TestMap.setShipType(Index)

    On Error Resume Next
    Maps(activemap).picPreview.setfocus

End Sub



Sub picradar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Move to the clicked position
    If Not loadedmaps(activemap) Then Exit Sub

    Dim tileX As Integer, tileY As Integer
  
    With Maps(activemap)
  
        If .UsingRadarFullMap Then
            tileX = X * (1024 / picradar.width)
            tileY = Y * (1024 / picradar.height)
        Else
            tileX = .radar_left + X
            tileY = .radar_top + Y
        End If
        
        If tileX < 0 Then tileX = 0
        If tileX >= MAPW Then tileX = MAPW - 1
        If tileY < 0 Then tileY = 0
        If tileY >= MAPH Then tileY = MAPH - 1
        
        
        'Change the view to the part where clicked, so that point on the radar
        'is the centered tile in the preview
        'SharedVar.MouseDown = True
        If Button = vbLeftButton Then
        
            'change the value of the scrollbars so that the point
            'where clicked on the radar, is in center of the screen
        
            Call .SetFocusAt(tileX, tileY, .picPreview.width \ 2, .picPreview.height \ 2, True)
        
        Else
        
            frmGoto.Xcoord = tileX
            frmGoto.Ycoord = tileY
        
            SetStretchBltMode frmGoto.picmap.hDC, HALFTONE
            'Call StretchBlt(frmGoto.picmap.hdc, 0, 0, frmGoto.picmap.width, frmGoto.picmap.height, Maps(activemap).pic1024.hdc, 0, 0, 1024, 1024, vbSrcCopy)
            Call .cpic1024.stretchToDC(frmGoto.picmap.hDC, 0, 0, frmGoto.picmap.width, frmGoto.picmap.height, 0, 0, 1024, 1024, vbSrcCopy)
            
            frmGoto.picmap.Refresh
        
            frmGoto.show vbModal, frmGeneral
        End If
    End With
End Sub

Sub picradar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show the position where we hover over
    If Not loadedmaps(activemap) Then Exit Sub

    Dim posX As Integer, posY As Integer

    If Maps(activemap).UsingRadarFullMap Then
        posX = X * (1024 / picradar.width)
        posY = Y * (1024 / picradar.height)
        If Button = vbLeftButton Then
            Call picradar_MouseDown(Button, Shift, X, Y)
        End If
    Else
        posX = Maps(activemap).radar_left + X
        posY = Maps(activemap).radar_top + Y
    End If

    If posX >= 0 And posY >= 0 And posX < 1024 And posY < 1024 Then
        If frmGeneral.lblposition.Caption <> "X= " & posX & " - Y= " & posY & " (" & Chr(65 + Int(posX / Int(1024 / 20))) & 1 + Int(posY / Int(1024 / 20)) & ")" & " (Radar)" Then
            'only update label when it has actually changed (to prevent flickering)
            frmGeneral.lblposition.Caption = "X= " & posX & " - Y= " & posY & " (" & Chr(65 + Int(posX / Int(1024 / 20))) & 1 + Int(posY / Int(1024 / 20)) & ")" & " (Radar)"
        End If
    Else
        If frmGeneral.lblposition.Caption <> "Outside map" Then
            'only update label when it has actually changed (to prevent flickering)
            frmGeneral.lblposition.Caption = "Outside map"
        End If
    End If

    'If Button Then
    '    Call picradar_MouseDown(Button, Shift, x, y)
    'End If
End Sub


Private Sub picRadarPopup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not loadedmaps(activemap) Then Exit Sub
    PopupRadar
End Sub





Private Sub picsmalltilepreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not loadedmaps(activemap) Then Exit Sub
    PopupTileset
End Sub






Private Sub cmdToggleRightPanel_Click()
'Show and hides the appropiate stuff when toggling the right panel
    Dim i As Integer
    
    If cmdToggleRightPanel.Caption = "->" Then
        dontResizeRadar = True
        cmdToggleRightPanel.Caption = "<-"
        
        cTileset.visible = False
        
        picradar.visible = False
        Line1.visible = False
        lblpositionA.visible = False
        lblFromA.visible = False
        lblToA.visible = False
        pictilesetlarge.visible = False
                
        picRightBar.width = 405
        picsmalltilepreview.visible = True
        picRadarPopup.visible = True
        'If loadedmaps(activemap) Then Call MDIForm_Resize
        
    Else
        dontResizeRadar = False
        picsmalltilepreview.visible = False
        picRadarPopup.visible = False
        cmdToggleRightPanel.Caption = "->"
        picRightBar.width = 4600
        
        cTileset.visible = True
        
        picradar.visible = True
        Line1.visible = True
        lblpositionA.visible = True
        lblFromA.visible = True
        lblToA.visible = True
        pictilesetlarge.visible = True
        
       
        'If loadedmaps(activemap) Then Call MDIForm_Resize
    End If
  

End Sub






Private Sub sldAirbDensity_Change()
    txtAirBrDensity.Text = sldAirbDensity.value
    Call SetSetting("AirbDensity", sldAirbDensity.value)
End Sub

Private Sub sldAirbDensity_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call sldAirbDensity_Change
End Sub

Private Sub sldAirbDensity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then Call sldAirbDensity_Change
End Sub

Private Sub sldAirbSize_Change()
    txtAirBrSize.Text = sldAirbSize.value
    Call SetSetting("AirbSize", sldAirbSize.value)
End Sub
Private Sub sldAirbSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call sldAirbSize_Change
End Sub
Private Sub sldAirbSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then Call sldAirbSize_Change
End Sub


Sub UpdateTileTextData(KeyAscii As Integer)
    TileTextData.setfocus
    Call TileTextData_KeyPress(KeyAscii)
    Maps(activemap).picPreview.setfocus
End Sub


Private Sub tbMaps_Click()
'10
End Sub

Private Sub tbMaps_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuToolbars
    Else
        Call mnuMaps_Click(val(tbMaps.SelectedItem.Key))
    End If
End Sub

Private Sub tlbTileset_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
    Case "EditTileset"
        Call mnuEditTileset_Click
    
    Case "ImportTileset"
        Call mnuImportTileset_Click
        
    Case "ExportTileset"
        Call mnuExportTileset_Click
        
    Case "DiscardTileset"
        Call mnudiscardtileset_Click
    
    Case "EditWalltiles"
        Call mnuWalltiles_Click
    
    Case "EditLVZ"
        Call mnuManageLVZ_Click
        
    End Select
End Sub





Private Sub TileTextData_KeyPress(KeyAscii As Integer)
'''''This needs to stay for TileText to work
End Sub

Private Sub tlbTabs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuToolbars
    End If
End Sub

Private Sub tlbToolOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuToolbars
    End If
End Sub



Private Sub ToolbarLeft_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not loadedmaps(activemap) Then Exit Sub

    'Sets the tool you selected
    Call SetCurrentTool(Button.Index)

    If Button.Key = "CustomShape" Then
        Call PopupMenu(customShapeMenu)
    End If

    If Button.Key = "TestMap" And Not Maps(activemap).TestMap.isRunning And _
       curtool = T_TestMap Then
        Call Maps(activemap).TestMap.StartRun
    End If


  'set focus back to the the activemap
  On Error Resume Next
  
  'The activemap could have been unloaded by now during a DoEvents
  If loadedmaps(activemap) Then Maps(activemap).picPreview.setfocus
End Sub

Private Sub ToolbarTop_ButtonClick(ByVal Button As MSComctlLib.Button)
'Do the appropiate stuff with the buttons
    Select Case Button.Key
    Case "New"
        Call mnuNew_Click

    Case "Open"
        Call mnuOpen_Click

    Case "Save"
        Call mnuSave_Click

    Case "Cut"
        Call mnuCut_Click

    Case "Copy"
        Call mnuCopy_Click

    Case "Paste"
        Call mnuPaste_Click

    Case "Undo"
        Call mnuUndo_Click

    Case "Redo"
        Call mnuRedo_Click

    Case "Grid"
        Call mnuGrid_Click

    Case "TileNr"
        Call mnuTileNR_Click

    Case "WallTiles"
        Call mnuWalltiles_Click

    Case "Mirror"
        Call mnuMirror_Click

    Case "Flip"
        Call mnuFlip_Click

    Case "Rotate"
        Call mnuRotate_Click

    Case "Replace"
        Call mnuReplace_Click

    Case "TextToMap"
        Call mnuTextToMap_Click

    Case "PicToMap"
        Call mnuPTM_Click
    
    Case "PasteType"
        PopupMenu mnuDrawMode
        
    Case "ZoomIn"
        Call ExecuteZoom(False)
    Case "ZoomOut"
        Call ExecuteZoom(True)


    Case "ShowRegions"
        Call mnuShowRegions_Click

    Case "ShowLVZ"
        Call mnuShowLVZ_Click
        
    Case "EditELVL"
        Call mnuElvl_Click

    Case "ManageLVZ"
        Call mnuManageLVZ_Click
    End Select
End Sub

Private Sub frmTool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuToolbars
    End If
End Sub

Private Sub Toolbarleft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DoEvents
    If Button = vbRightButton Then
        PopupMenu mnuToolbars
    ElseIf Button = vbLeftButton Then
        If Y >= Toolbarleft.Buttons("CustomShape").Top _
           And Y < Toolbarleft.Buttons("CustomShape").Top + Toolbarleft.Buttons("CustomShape").height Then
            'clicked on customshape
            Call ToolbarLeft_ButtonClick(Toolbarleft.Buttons("CustomShape"))
        End If

    End If
End Sub

Private Sub toolbartop_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
    Case "PasteType"
        PopupMenu mnuDrawMode
    
    End Select
End Sub

Private Sub ToolbarTop_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    
    Case "PasteNormal"
        Call mnuNormalPaste_Click

    Case "PasteUnder"
        Call mnuPasteUnder_Click

    Case "PasteTransparent"
        Call mnuTransparentPaste_Click
        
    End Select
    
End Sub

Private Sub toolbartop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuToolbars
    End If
End Sub

Private Sub toolSize_Change(Index As Integer)
    Call SetSetting("ToolSize" & ToolName(Index), toolSize(Index).value)
End Sub

Private Sub optToolRound_Click(Index As Integer)
    Call SetSetting("ToolTip" & ToolName(Index), "0")
End Sub
Private Sub optToolSquare_Click(Index As Integer)
    Call SetSetting("ToolTip" & ToolName(Index), "1")
End Sub

Private Sub toolStep_Change(Index As Integer)
    Call SetSetting("ToolStep" & ToolName(Index), toolStep(Index).value)
End Sub

Private Sub txtAirBrsize_Change()
    Call removeDisallowedCharacters(txtAirBrSize, 1, MAX_AIRBR_SIZE, True)
    sldAirbSize.value = val(txtAirBrSize.Text)
End Sub
Private Sub txtAirBrsize_Click()
    Call toggleLockToolTextBox(txtAirBrSize)
End Sub
Private Sub txtAirBrsize_LostFocus()
    Call toggleLockToolTextBox(txtAirBrSize, False)
End Sub

Private Sub txtAirBrDensity_Change()
    Call removeDisallowedCharacters(txtAirBrDensity, 1, MAX_AIRBR_DENSITY, True)
    sldAirbDensity.value = val(txtAirBrDensity.Text)
End Sub
Private Sub txtAirBrDensity_Click()
    Call toggleLockToolTextBox(txtAirBrDensity)
End Sub
Private Sub txtAirBrDensity_LostFocus()
    Call toggleLockToolTextBox(txtAirBrDensity, False)
End Sub



' ---------------------------------------------------------------
' ---------------------------------------------------------------
' ---------------------- DEFINED METHODS ------------------------
' ---------------------------------------------------------------
' ---------------------------------------------------------------

Sub ExecuteCount()
'Counts the tiles in the active map
    If Not loadedmaps(activemap) Then Exit Sub
    
    Load frmCount
    
    Call frmCount.setParent(Maps(activemap))
    Call frmCount.CountTiles
    
    frmCount.show vbModal, frmGeneral
    
End Sub

Sub ExecuteGoTo(X As Integer, Y As Integer)
    If Not loadedmaps(activemap) Then Exit Sub

    With Maps(activemap)
        If .TestMap.isRunning Then
            Call .TestMap.WarpShip(X, Y)
        Else
            Call .SetFocusAt(X, Y, .picPreview.width \ 2, .picPreview.height \ 2, True)
        End If
    End With
    '    Call picradar_MouseDown(vbLeftButton, 0, X - Maps(activemap).radar_left, Y - Maps(activemap).radar_top)
    '    SharedVar.MouseDown = False
End Sub

Sub ExecuteTextToMap(Text() As Integer, width As Integer, height As Integer)
'Paste the Text To Map from the clipboard
    If Not loadedmaps(activemap) Then Exit Sub
    Call Maps(activemap).sel.TextToMap(Text, width, height)
End Sub

'Samapico: This is now useless, frmPicToMap now has a reference to
'          the parent map, so it can call PicToMap() directly
'Sub ExecutePicToMap(Pic() As Integer, width As Integer, height As Integer)
''Paste the Picture To Map from the clipboard
'    If Not loadedmaps(activemap) Then Exit Sub
'    Call Maps(activemap).sel.PicToMap(Pic, width, height)
'End Sub

Sub ExecuteReplace(mode As replaceenum, tilesrc As Integer, tiledest As Integer, inselection As Boolean, RedrawWalltiles As Boolean)
'Switch/Replace tiles in the active map
    If Not loadedmaps(activemap) Then Exit Sub
    frmGeneral.IsBusy("frmGeneral.ExecuteReplace") = True
    If Not inselection Then
        'not in selection, use the normal map function
        Select Case mode
        Case switchleftright
            Call Maps(activemap).SwitchOrReplace(tilesrc, tiledest, False, RedrawWalltiles)
        Case replaceleftright
            Call Maps(activemap).SwitchOrReplace(tilesrc, tiledest, True, RedrawWalltiles)
        End Select
    Else
        'in selection, do the switch/replace in sel defined method
        Select Case mode
        Case switchleftright
            Call Maps(activemap).sel.SwitchOrReplace(tilesrc, tiledest, False, RedrawWalltiles)
        Case replaceleftright
            Call Maps(activemap).sel.SwitchOrReplace(tilesrc, tiledest, True, RedrawWalltiles)
        End Select
    End If
    frmGeneral.IsBusy("frmGeneral.ExecuteReplace") = False
End Sub

Sub ExecuteMirror()
'Mirrors the selection on the active map
    If Not loadedmaps(activemap) Then Exit Sub
    
    If Maps(activemap).sel.hasAlreadySelectedParts Then
        'mirror the selection
        frmGeneral.IsBusy("frmGeneral.ExecuteMirror") = True
        
        Maps(activemap).undoredo.ResetRedo
        Dim undoch As New Changes

        Call Maps(activemap).sel.Mirror(undoch)

        Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTION_MIRROR)
        
        frmGeneral.IsBusy("frmGeneral.ExecuteMirror") = False
    Else
        'not appending, so no valid selection
    End If
    
End Sub

Sub ExecuteFlip()
'Flips the selection on the active map
    If Not loadedmaps(activemap) Then Exit Sub
    
    If Maps(activemap).sel.hasAlreadySelectedParts Then
        'flip the selection
        frmGeneral.IsBusy("frmGeneral.ExecuteFlip") = True
        
        Maps(activemap).undoredo.ResetRedo
        Dim undoch As New Changes

        Call Maps(activemap).sel.Flip(undoch)

        Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTION_FLIP)
        
        frmGeneral.IsBusy("frmGeneral.ExecuteFlip") = False
    Else
        'not appending, so no valid selection
    End If
End Sub

Sub ExecuteRotate(direction As Integer, Optional angle As Double)
'Rotates the selection given the angle
    If Not loadedmaps(activemap) Then Exit Sub
    
    Dim undoch As Changes
    frmGeneral.IsBusy("frmGeneral.ExecuteRotate") = True
    
    If direction = 1 Then
        Maps(activemap).undoredo.ResetRedo
        Set undoch = New Changes
        Call Maps(activemap).sel.RotateCW(undoch)
        Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTION_ROTATE90)

    ElseIf direction = 2 Then
        Maps(activemap).undoredo.ResetRedo
        Set undoch = New Changes
        Call Maps(activemap).sel.Rotate180(undoch)
        Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTION_ROTATE180)

    ElseIf direction = 3 Then
        Maps(activemap).undoredo.ResetRedo
        Set undoch = New Changes
        Call Maps(activemap).sel.RotateCCW(undoch)
        Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTION_ROTATE270)

    ElseIf direction = 4 Then
        Maps(activemap).undoredo.ResetRedo
        Set undoch = New Changes
        Call Maps(activemap).sel.RotateFree(angle, undoch)
        Call Maps(activemap).undoredo.AddToUndo(undoch, UNDO_SELECTION_ROTATEFREE)
    End If
    
    frmGeneral.IsBusy("frmGeneral.ExecuteRotate") = False
End Sub

Sub ExecuteZoom(out As Boolean)
'Does a zoom on the active map, used for zooming with toolbar buttons

    If Not loadedmaps(activemap) Then Exit Sub

    frmGeneral.IsBusy("frmGeneral.ExecuteZoom") = True
    
    'use center tile for zooming in or out
    If Not out Then
        Call Maps(activemap).magnifier.ZoomIn( _
             (Int(((Maps(activemap).hScr.value + Maps(activemap).picPreview.width / 2)) / (TILEW * Maps(activemap).magnifier.zoom))), _
             (Int(((Maps(activemap).vScr.value + Maps(activemap).picPreview.height / 2)) / (TILEW * Maps(activemap).magnifier.zoom))))
    Else
        Call Maps(activemap).magnifier.ZoomOut( _
             (Int(((Maps(activemap).hScr.value + Maps(activemap).picPreview.width / 2)) / (TILEW * Maps(activemap).magnifier.zoom))), _
             (Int(((Maps(activemap).vScr.value + Maps(activemap).picPreview.height / 2)) / (TILEW * Maps(activemap).magnifier.zoom))))
    End If

    frmGeneral.IsBusy("frmGeneral.ExecuteZoom") = False
End Sub


Friend Sub ExecuteZoomFocus(out As Boolean, ByRef mousepos As POINTAPI)
'Does a zoom on the active map focusing on current cursor position, used for zooming with mousewheel
    Dim tileX As Integer
    Dim tileY As Integer

    If Not loadedmaps(activemap) Then Exit Sub
    'curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    'curtiley = (parent.Vscr.value + Y) \ parent.currenttilew
    ScreenToClient Maps(activemap).picPreview.hWnd, mousepos
    
    Dim offsetX As Integer
    Dim offsetY As Integer

    frmGeneral.IsBusy("frmGeneral.ExecuteZoomFocus") = True
    
'30        offsetX = (Me.Left + Maps(activemap).Left + Toolbarleft.width) / Screen.TwipsPerPixelX
'40        offsetY = (Me.Top + Maps(activemap).Top + IIf(toolbartop.visible, toolbartop.height, 0) + IIf(tlbToolOptions.visible, tlbToolOptions.height, 0)) / Screen.TwipsPerPixelY
'
'50        tileX = (Maps(activemap).Hscr.Value + mouseX - offsetX) \ Maps(activemap).currenttilew
'60        tileY = (Maps(activemap).Vscr.Value + mouseY - offsetY) \ Maps(activemap).currenttilew
    tileX = (Maps(activemap).hScr.value + mousepos.X) \ Maps(activemap).currenttilew
    tileY = (Maps(activemap).vScr.value + mousepos.Y) \ Maps(activemap).currenttilew

    'use center tile for zooming in or out
    If Not out Then
        Call Maps(activemap).magnifier.ZoomIn(tileX, tileY)
    Else
        Call Maps(activemap).magnifier.ZoomOut(tileX, tileY)
    End If

    frmGeneral.IsBusy("frmGeneral.ExecuteZoomFocus") = False
End Sub

Sub ExecuteUpdate(Optional UpdateSettings As Boolean = False)
'updates the current map
    If Not loadedmaps(activemap) Then Exit Sub

    If UpdateSettings Then Call Maps(activemap).StoreSettings

    Call Maps(activemap).UpdateLevel
End Sub

Private Function DefaultWindowState() As FormWindowStateConstants
'Returns the windowstate to which a new or opened map should be set to by default

    If Not loadedmaps(activemap) Then
        'If no maps are loaded yet, default to maximized
        DefaultWindowState = vbMaximized
    Else
        If Maps(activemap).windowstate = vbMinimized Then
            DefaultWindowState = vbNormal
        Else
            DefaultWindowState = Maps(activemap).windowstate
        End If
    End If
End Function

Sub NewMap()
'Creates a new map if there is still space available

'get a free slot for the map

    Dim freemap As Integer
    Dim windowstate As FormWindowStateConstants
  
  If Not bDEBUG Then
    On Error GoTo NewMap_Error
  End If
    frmGeneral.IsBusy("frmGeneral.NewMap") = True
    
    freemap = getFreeMapSlot()

    If freemap = -1 Then
        frmGeneral.IsBusy("frmGeneral.NewMap") = False
        'no free slots anymore, don't create a new map
        MessageBox "Error: Maximum number of maps are already open", vbExclamation
        Exit Sub
    End If

    windowstate = DefaultWindowState

    'load a new mapmap
    Set Maps(freemap) = New frmMain
    Load Maps(freemap)
    
    Maps(freemap).windowstate = windowstate
    
    loadedmaps(freemap) = True
    

    Maps(freemap).id = freemap
    'set current loaded map on the active one, because a new
    'form that is newly loaded is always the active one
    activemap = freemap

    'check if the maps name is already used (untitled that is)
    'and if so, keep counting until we have a map that isn't called
    'Untitled<i>
    Dim str As String
    Dim count As Integer
    Dim i As Integer
    str = "Untitled"
retry:
    str = "Untitled " & count + 1
    i = 0
    For i = 0 To 9
        If loadedmaps(i) And i <> freemap Then
            If Maps(i).Caption = str Then
                count = count + 1
                GoTo retry
            End If
        End If
    Next
    'ok, we've got the default appended name
    Maps(freemap).Caption = str



'        Dim t As MSComctlLib.Tab
'
'        Set t = tbMaps.Tabs.add(activemap, activemap & "_", Maps(activemap).Caption)
'        t.selected = True
  
    'call the new map function in the map itself
    Call Maps(freemap).NewMap

    'update icon for that new picpreview
    Call SetCurrentTool(curtool)
'290       DoEvents
    'everything is loaded, show it to the user
    mnuMaps(activemap).visible = True
    mnuMaps(activemap).Caption = Maps(activemap).Caption


  
  
    Maps(freemap).picPreview.Enabled = True
    
    Call cTileset.setParent(Maps(freemap))
    Call Maps(freemap).tileset.SelectTiles(vbLeftButton, 1, 1, 1, True)
    Call Maps(freemap).tileset.SelectTiles(vbRightButton, 2, 1, 1, False)

    Call UpdateToolBarButtons
    Call UpdateToolToolbar
    
    frmGeneral.IsBusy("frmGeneral.NewMap") = False
    
    On Error GoTo 0
    Exit Sub

NewMap_Error:
    frmGeneral.IsBusy("frmGeneral.NewMap") = False
    HandleError Err, "frmGeneral.NewMap"
End Sub



Function GetIndexOfMap(path As String) As Integer

    Dim i As Integer
    For i = 0 To 9
        If loadedmaps(i) Then
            If Maps(i).activeFile = path Then
                GetIndexOfMap = i
                Exit Function
            End If
        End If
    Next
    GetIndexOfMap = -1
End Function


Function OpenMap(path As String) As Integer
'Opens a map, and returns the map id

    Dim lvlpath As String
    'get a free slot for the map
    Dim freemap As Integer

    'no free slots
    On Error GoTo OpenMap_Error
    
    OpenMap = -1
    
    freemap = getFreeMapSlot()
    If freemap = -1 Then
        MessageBox "Error: Maximum number of maps are already open", vbExclamation
        'return -1 so that when we are opening multiple files,
        'we wont get 10 different msgboxes that say that the
        'maximum number of maps are open
        OpenMap = freemap
        Exit Function
    End If

'          Dim f As Integer
'80        f = FreeFile

'          Dim bm As Integer
'          Dim size As Long
'          Dim b() As Byte

    On Error GoTo errh

    'if no path is given, show the common dialog, else open the
    'path (is used when dropping on the mdi form)
    If path = "" Then
      cd.InitDir = GetLastDialogPath("OpenMap")
      
        cd.DialogTitle = "Select an lvl to open"
        cd.flags = cdlOFNHideReadOnly
        cd.Filter = "Level files (*.lvl, *.elvl)|*.lvl; *.bak; *.elvl|All files (*.*)|*.*"
        cd.ShowOpen
        lvlpath = cd.filename
      Call SetLastDialogPath("OpenMap", GetPathTo(lvlpath))
      cd.InitDir = ""
    Else
        lvlpath = path
    End If
    
    frmGeneral.IsBusy("frmGeneral.OpenMap") = True
    
    'lvlpath holds now the correct path for the map to open

    'check if the map isn't already open
    
    If GetIndexOfMap(lvlpath) <> -1 Then
         frmGeneral.IsBusy("frmGeneral.OpenMap") = False
         MessageBox "The file '" & lvlpath & "' is already opened!", vbExclamation
         Exit Function
    End If

    
    If Not FileExists(lvlpath) Then
        frmGeneral.IsBusy("frmGeneral.OpenMap") = False
        MessageBox "File '" & lvlpath & "' was not found.", vbOKOnly + vbExclamation
        Exit Function
    End If
    
'    If GetExtension(lvlpath) <> "lvl" And GetExtension(lvlpath) <> "bak" And GetExtension(lvlpath) <> "elvl" Then
'        If messagebox("Specified file does not have a .lvl extension and might not be a valid map file, do you wish to load it anyway?", vbYesNo + vbCritical, "Invalid file type") = vbNo Then
'            frmGeneral.IsBusy("frmGeneral.OpenMap") = False
'            Exit Function
'        End If
'    End If
    ShowProgress "Loading map", 1624
    
    Call UpdateProgressLabel("Loading form...")
    
    If loadedmaps(activemap) Then
        If Maps(activemap).activeFile = "" And Not Maps(activemap).mapchanged Then
            'we have an empty unchanged Untitled map opened, close it
            Call DestroyMap(activemap)
        End If
    End If
    
    freemap = getFreeMapSlot()
    
    'ok we will open a map
    'prepare a map
    On Error Resume Next 'gah, when a modal form is open and we drag drop open or anything, we get an error here :/
    
    Dim windowstate As FormWindowStateConstants
    
    windowstate = DefaultWindowState
    
    Set Maps(freemap) = New frmMain
    
    Maps(freemap).show
    
    Maps(freemap).windowstate = windowstate
    
    On Error GoTo OpenMap_Error
    loadedmaps(freemap) = True
    Maps(freemap).id = freemap
    activemap = freemap

    'call the open map function of the map itself,
    'and pass the filepointer and path
    Maps(freemap).Caption = GetFileTitle(lvlpath)

    Call Maps(freemap).OpenMap(lvlpath)
'510       DoEvents
    'update icon for that new picpreview
    frmGeneral.SetCurrentTool (curtool)
    
    'update the window of the map
    Maps(freemap).UpdateLevel

    Maps(freemap).picPreview.Enabled = True

    'set the windows menu
    mnuMaps(activemap).visible = True
    mnuMaps(activemap).Caption = Maps(activemap).Caption

    ' if the map is succesfully opened, return the map id
    ' else return -1
    OpenMap = freemap

    Call cTileset.setParent(Maps(activemap))
    Call Maps(activemap).tileset.SelectTiles(vbLeftButton, 1, 1, 1, True)
    Call Maps(activemap).tileset.SelectTiles(vbRightButton, 2, 1, 1, True)
    
    Call UpdateToolBarButtons

    frmGeneral.IsBusy("frmGeneral.OpenMap") = False
    
    Exit Function

errh:
    If Err = cdlCancel Then
        frmGeneral.IsBusy("frmGeneral.OpenMap") = False
        Exit Function
    Else
        frmGeneral.IsBusy("frmGeneral.OpenMap") = False
        HandleError Err, "frmGeneral.OpenMap:errh"
    End If

    On Error GoTo 0
    Exit Function

OpenMap_Error:

    If freemap <> -1 Then
        loadedmaps(freemap) = False
        activemap = 0
    End If
    
    frmGeneral.IsBusy("frmGeneral.OpenMap") = False
    HandleError Err, "frmGeneral.OpenMap:OpenMap_Error " & path
End Function

Function SaveMap(Optional saveas As Boolean = True, Optional flags As saveFlags = SFdefault) As Boolean
'Save the active map
    On Error GoTo SaveMap_Error
    
    If Not loadedmaps(activemap) Then Exit Function
    
    Dim tick As Long
    tick = GetTickCount
    
    ' Check if the number of flags are less than 256
    ' if this check would be done with each setTile, then this would
    ' create way too much overhead
    
    Dim FlagsCount As Long
    FlagsCount = Maps(activemap).CountTile(170)
    If FlagsCount > 256 Then
        MessageBox "You have " & FlagsCount & " flags. You can't have more than 256!", vbExclamation
        SaveMap = False
        Exit Function
    End If

'   This warning apparently was more confusing than anything else for some people
'   If they used transparent or draw-under selection, they could easily have an invisible selection too
'    If Maps(activemap).sel.hasAlreadySelectedParts Then
'
'        MessageBox "You have an active selection in your map. The tiles in that selection will not be saved with your map. Save your map again after you unselected them.", vbExclamation
'    End If
    
    frmGeneral.IsBusy("frmGeneral.SaveMap") = True
    
    'if no path is present in the activeFile in the map (set when
    'saved or loaded) then show save as, also show save as if explicitly
    'asked for it
    If saveas Or Maps(activemap).activeFile = "" Then    'no active file
        On Error GoTo errh
      cd.InitDir = GetLastDialogPath("SaveMap")
      
        cd.DialogTitle = "Save lvl as..."
        'ask for overwrite
        cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        cd.filename = Maps(activemap).Caption
        cd.Filter = "Level file (*.lvl)|*.lvl|Extended level file (*.elvl)|*.elvl|All files (*.*)|*.*"
        cd.ShowSave
      Call SetLastDialogPath("SaveMap", GetPathTo(cd.filename))
      cd.InitDir = ""
      
        'don't overwrite the former path to the new file where special
        'data is lost
        If FlagIs(flags, SFsaveExtraTiles) Then
            'if the file is saved somewhere else, change the activefile
            'to that path
            If Maps(activemap).activeFile <> cd.filename Then
                Maps(activemap).activeFile = cd.filename
                Call AddRecent(cd.filename)
                Maps(activemap).Caption = cd.filetitle
            End If
        End If

        ShowProgress "Saving", 1524
        
        'call the save handling in the map itself
        Call Maps(activemap).SaveMap(cd.filename, flags)

    Else

        ShowProgress "Saving", 1524
        
        'no save as, just saving
        Call Maps(activemap).SaveMap(Maps(activemap).activeFile, flags)
    End If

    SaveMap = True
    
    Call UpdateToolBarButtons
    
    Label6.Caption = "Save: " & GetTickCount - tick
    
    frmGeneral.IsBusy("frmGeneral.SaveMap") = False
    Exit Function
errh:
    If Err = cdlCancel Then
        SaveMap = False
        frmGeneral.IsBusy("frmGeneral.SaveMap") = False
        Exit Function
    Else
        frmGeneral.IsBusy("frmGeneral.SaveMap") = False
        MessageBox Err & " " & Err.description, vbCritical
    End If

    On Error GoTo 0
    Exit Function

SaveMap_Error:
    frmGeneral.IsBusy("frmGeneral.SaveMap") = False
    HandleError Err, "frmGeneral.SaveMap"
End Function

Sub DestroyMapReference(id)
    
    
    Set Maps(id) = Nothing
    
    
End Sub


Sub DestroyMap(id)
'Closes the map

'remove the object reference
    On Error GoTo DestroyMap_Error

    If Not loadedmaps(id) Then
        Set Maps(id) = Nothing
        Exit Sub
    End If

    Dim Caption As String
    Caption = Maps(id).Caption
    
    frmGeneral.IsBusy("frmGeneral.DestroyMap") = True
    
    AddDebug "+++ Closing map " & Maps(id).Caption & " (ID: " & id & ")"

'80        Call Maps(id).ClearRevert
    
    'MAPS MUST BE UNLOADED...
    loadedmaps(id) = False
    
    Unload Maps(id)

    'Clear references to the map
    Call DestroyMapReference(id)

    'We know it's unloaded now
    
    
    
    
    'remove from the menu
    mnuMaps(id).visible = False
    mnuMaps(id).Caption = ""

'        tbMaps.Tabs.Remove (id)
  
    'Check if there are any more loaded maps
    Dim i As Integer
    For i = 0 To 9
        If loadedmaps(i) = True Then
            frmGeneral.IsBusy("frmGeneral.DestroyMap") = False
            Exit Sub
        End If
    Next

    'There are no more maps, if there were, we would have Exit Subbed
    'clear the tileset
    'TOREMOVE---REMOVED
'    pictileset.Cls
'    pictileset.Picture = LoadPicture("")
    Call cTileset.ClearTileset
    Call cTileset.setParent(Nothing)
    
    picradar.Cls
    pictilesetlarge.Cls
    picsmalltilepreview.Cls

    
    Call UpdateToolBarButtons
    Call UpdateToolToolbar
    
    frmGeneral.IsBusy("frmGeneral.DestroyMap") = False
    
    On Error GoTo 0
    Exit Sub

DestroyMap_Error:
    frmGeneral.IsBusy("frmGeneral.DestroyMap") = False
    
    HandleError Err, "frmGeneral.DestroyMap"
End Sub

Function getFreeMapSlot() As Integer
'Checks if there are any free spots to load a map

'if not return -1 else return the available slot
    Dim i As Integer
    For i = 0 To 9
        If loadedmaps(i) = False Then
            getFreeMapSlot = i
            Exit Function
        End If
    Next

    getFreeMapSlot = -1
End Function

Sub SaveMiniMap(p As String, dimension As Integer)
' Saves the minimap with the given size to the given path
    On Error GoTo SaveMiniMap_Error

    On Error GoTo errh
    Dim path As String
    If p = "" Then
        'no path is given, show dialog
      cd.InitDir = GetLastDialogPath("SaveMinimap")
      
        cd.Filter = "*.bmp|*.bmp"
        cd.DialogTitle = "Save Mini Map with size " & dimension
        'ask for overwrite
        cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        cd.filename = ""
        cd.ShowSave

        path = cd.filename
      Call SetLastDialogPath("SaveMinimap", path)
      cd.InitDir = ""
    Else
        path = p
    End If

    'init the temp dimensions
    pictemp.width = dimension
    pictemp.height = dimension

    Call AddDebug("Saving radar image at " & path & " with dimension: " & dimension)

    If dimension <> 1024 Then
        'stretch it onto the pictemp
        SetStretchBltMode pictemp.hDC, HALFTONE
        'StretchBlt picTemp.hdc, 0, 0, dimension, dimension, Maps(activemap).pic1024.hdc, 0, 0, 1024, 1024, vbSrcCopy
        Call Maps(activemap).cpic1024.stretchToDC(pictemp.hDC, 0, 0, dimension, dimension, 0, 0, 1024, 1024, vbSrcCopy)
    Else
        'BitBlt picTemp.hdc, 0, 0, dimension, dimension, Maps(activemap).pic1024.hdc, 0, 0, vbSrcCopy
        Call Maps(activemap).cpic1024.bltToDC(pictemp.hDC, 0, 0, dimension, dimension, 0, 0, vbSrcCopy)
    End If

    'apply to the picture
    pictemp.Picture = pictemp.Image

    'save it
    Call SavePicture(pictemp.Picture, path)

    'clean up
    pictemp.Cls
    pictemp.Picture = LoadPicture("")

    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    End If

    On Error GoTo 0
    Exit Sub

SaveMiniMap_Error:
    HandleError Err, "frmGeneral.SaveMiniMap"
End Sub



'Sub SetLeftSelection(tileNr As Integer, sizeX As Integer, sizeY As Integer, Optional setLastButton As Boolean = True)
''Sets the left selection of the tileset on the given tilenr
'    If tileNr = 0 Then Exit Sub
'    If tileNr > 256 And tileNr < 259 Then Exit Sub
'
'
'    If loadedmaps(activemap) Then
'        If tileNr >= 259 Then
'            If Maps(activemap).walltiles.isValidSet(tileNr - 259) Then
'                Maps(activemap).useWallTileLeft = True
'                Maps(activemap).tline.usewalltile = True
'                Maps(activemap).curWallTileLeft = tileNr - 259
'                Maps(activemap).tilesetleft = Maps(activemap).walltiles.getWallTile(Maps(activemap).curWallTileLeft, 0)
'                Maps(activemap).walltiles.curwall = Maps(activemap).curWallTileLeft
'                Maps(activemap).multTileLeftx = 1
'                Maps(activemap).multTileLefty = 1
'            Else
'                Exit Sub
'            End If
'        Else
'            'put the left tile into the active map
'            Maps(activemap).tilesetleft = tileNr Mod 256
'            Maps(activemap).multTileLeftx = sizeX
'            Maps(activemap).multTileLefty = sizeY
'
'            Maps(activemap).useWallTileLeft = False
'            Maps(activemap).tline.usewalltile = False
'        End If
'    End If
'
'    'move the shape
'    leftsel.Left = ((tileNr - 1) Mod 19) * 16
'    leftsel.Top = (Int((tileNr - 1) / 19)) * 16
'
'    leftsel.width = TileW * sizeX
'    leftsel.height = TileW * sizeY
'
'
'    'if the left and right tiles are the same, show a white shape to
'    'indicate they are the same
'    If rightsel.Left = leftsel.Left And rightsel.Top = leftsel.Top And rightsel.width = leftsel.width And rightsel.height = leftsel.height Then
'        rightsel.BorderColor = vbWhite
'    Else
'        rightsel.BorderColor = vbYellow
'    End If
'
'    If setLastButton Then Maps(activemap).lastButton = vbLeftButton
'
'    'Update the tileset preview
'    Call UpdatePreview
'End Sub

'Sub SetRightSelection(tileNr As Integer, sizeX As Integer, sizeY As Integer, Optional setLastButton As Boolean = True)
''Sets the right selection of the tileset on the given tilenr
'    If tileNr = 0 Then Exit Sub
'    If tileNr > 256 And tileNr < 259 Then Exit Sub
'
'
'
'    If loadedmaps(activemap) Then
'        If tileNr >= 259 Then
'            If Maps(activemap).walltiles.isValidSet(tileNr - 259) Then
'                Maps(activemap).useWallTileRight = True
'                Maps(activemap).tline.usewalltile = True
'                Maps(activemap).curWallTileRight = tileNr - 259
'                Maps(activemap).tilesetright = Maps(activemap).walltiles.getWallTile(Maps(activemap).curWallTileRight, 0)
'                Maps(activemap).walltiles.curwall = Maps(activemap).curWallTileRight
'                Maps(activemap).multTileRightx = 1
'                Maps(activemap).multTileRighty = 1
'            Else
'                Exit Sub
'            End If
'        Else
'            'put the right tile into the active map
'            Maps(activemap).tilesetright = tileNr Mod 256
'            Maps(activemap).multTileRightx = sizeX
'            Maps(activemap).multTileRighty = sizeY
'
'            Maps(activemap).useWallTileRight = False
'            Maps(activemap).tline.usewalltile = False
'        End If
'    End If
'
'    'move the shape
'    rightsel.Left = ((tileNr - 1) Mod 19) * 16
'    rightsel.Top = (Int((tileNr - 1) / 19)) * 16
'
'    rightsel.width = TileW * sizeX
'    rightsel.height = TileW * sizeY
'
'    'if the left and right tiles are the same, show a white shape to
'    'indicate they are the same
'    If rightsel.Left = leftsel.Left And rightsel.Top = leftsel.Top And rightsel.width = leftsel.width And rightsel.height = leftsel.height Then
'        rightsel.BorderColor = vbWhite
'    Else
'        rightsel.BorderColor = vbYellow
'    End If
'
'    If setLastButton Then Maps(activemap).lastButton = vbRightButton
'
'    'Update the tileset preview
'    Call UpdatePreview
'End Sub

Private Sub SetCurrentCursor(ByVal tool As toolenum)
    Dim i As Integer
    Dim newIcon As Long
    Dim mPic As StdPicture
    Dim cursor As Integer

    'special cases

    If tool = T_Region Then
        If optRegionSel(REGION_MAGICWAND).value = True Then
            tool = T_magicwand
        Else
            tool = T_selection
        End If
    End If

    
    
    newIcon = CreateIcon32x32(frmGeneral.imlToolbarIcons.ListImages( _
                              frmGeneral.Toolbarleft.Buttons(tool).Image _
                              ).Picture.Handle)
                              
    If newIcon <> 0 Then Set mPic = HandleToPicture(newIcon, False) Else Set mPic = frmGeneral.imlToolbarIcons.ListImages(frmGeneral.Toolbarleft.Buttons(tool).Image).Picture

    ' use a cross + instead of the generated mouseicon for:
    ' line, rectange, ellipse, special line, filled ellipse and filled rectangle
    ' use the Ibeam for Tiletext
    If tool = T_tiletext Then
        cursor = vbIbeam

    ElseIf tool <> T_line And _
           tool <> T_rectangle And _
           tool <> T_selection And _
           tool <> T_ellipse And _
           tool <> t_spline And _
           tool <> T_filledellipse And _
           tool <> T_filledrectangle And _
           tool <> T_customshape And _
           tool <> T_Region And _
           tool <> T_TestMap And _
           tool <> T_lvz And _
           tool <> T_freehandselection Then
        'put all the mousepointers of all open maps to custom
        cursor = vbCustom

    ElseIf tool = T_TestMap Or tool = T_lvz Then
        cursor = vbDefault
    Else
        'put all the mousepointers of all open maps to cross
        cursor = vbCrosshair
    End If

    'apply cursor to all loaded maps
    For i = 0 To 9
        If loadedmaps(i) = True Then
            Maps(i).picPreview.MousePointer = cursor

            If cursor = vbCustom And mPic.Type = vbPicTypeIcon Then
                ' set the mouseicon property, unless the listimage item was a bitmap
                Set Maps(i).picPreview.MouseIcon = mPic
            End If

        End If
    Next


    Set mPic = Nothing

End Sub


Sub SetCurrentTool(tool As toolenum, Optional setToolbars As Boolean = True)
'Sets the current tool and update the mouseicon and such
' On Error GoTo SetCurrentTool_Error
    Dim oldtool As toolenum
    oldtool = curtool



    If loadedmaps(activemap) Then

        
        If SharedVar.splineInProgress Then
            Call Maps(activemap).SPline.MouseDown(vbRightButton, 0, 0)
        End If
        
        If Not Maps(activemap).magnifier.UsingPixels Then
            Call Maps(activemap).UpdatePreview
        End If

        If Maps(activemap).TileText.isActive = True And tool <> T_hand And _
           tool <> T_magnifier And _
           tool <> T_tiletext Then
            Call Maps(activemap).TileText.StopTyping
        End If

        If Maps(activemap).TestMap.isRunning And _
           tool <> T_TestMap Then
            Call Maps(activemap).TestMap.StopRun
        End If
    End If

    curtool = tool
    
    
    Call SetCurrentCursor(tool)

    'let the current tool be pressed in the toolbar
    Dim tlbbutton As Button
    For Each tlbbutton In Toolbarleft.Buttons
        If tlbbutton.Index = curtool Then
            tlbbutton.value = tbrPressed
        Else
            tlbbutton.value = tbrUnpressed
        End If
    Next
    

    'hide displayed coordinates
    frmGeneral.lblToA.visible = False
    frmGeneral.lblFromA.visible = False
    frmGeneral.lblTo.visible = False
    frmGeneral.lblFrom.visible = False

    If setToolbars Then
        Call setToolToolbar(curtool)
        Call UpdateToolToolbar
    End If

    If loadedmaps(activemap) Then
    If curtool = T_Region Or oldtool = T_Region Then
        'force redraw
        Call Maps(activemap).UpdateLevel
    
    ElseIf curtool = t_spline Then
        Maps(activemap).tileset.lastButton = vbLeftButton
    End If
    End If
    On Error GoTo 0
    Exit Sub

SetCurrentTool_Error:
    HandleError Err, "frmGeneral.SetCurrentTool"
End Sub

Sub setToolToolbar(tool As toolenum)
'Sets the toolbar to the given tool
    Dim i As Integer
    For i = frmTool.LBound To frmTool.UBound
        If toolOptionExist(i + 1) Then
            frmTool(i).visible = False
        End If
    Next


    If CBool(GetSetting("PinToolOptions", "1")) Or (GetSetting("ShowToolbarToolOptions", "1") And toolOptionExist(tool)) Then
        tlbToolOptions.visible = True
        If toolOptionExist(tool) Then frmTool(tool - 1).visible = True
    Else
        tlbToolOptions.visible = False
        If toolOptionExist(tool) Then
            frmTool(tool - 1).visible = False
        End If
    End If

    If tool = T_customshape Then
        For i = frmCustomShape.LBound To frmCustomShape.UBound
            If customShapeOptionExist(i + 0) Then
                frmCustomShape(i).visible = False
            End If
        Next
        frmCustomShape(curCustomShape).visible = True
    End If

End Sub

Sub ImportTileset()
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
        If Not hasLVLaTileset(cd.filename) Then
            'load default tileset
            If Maps(activemap).tilesetloaded = False Then
                Call Maps(activemap).InitTileset("")
                Exit Sub
            Else
                AddDebug "ImportTileset, No bitmap data found in " & cd.filename
                MessageBox "No tileset was found in " & cd.filetitle & ", if you wish to use the default tileset for your map, use Discard Tileset from the File menu", vbOKOnly + vbExclamation, "No tileset found"
                Exit Sub
            End If
        End If
    End If
    'imports the given tileset
    Call Maps(activemap).ImportTileset(cd.filename)

    Exit Sub
errh:
    'if something goes wrong, use the default tileset
    If Err = cdlCancel Then
        If Maps(activemap).tilesetloaded = False Then
            Call Maps(activemap).InitTileset("")
        Else
            Exit Sub
        End If
    Else
        MessageBox Err & " " & Err.description, vbCritical
    End If

    On Error GoTo 0
    Exit Sub

ImportTileset_Error:
    HandleError Err, "frmGeneral.ImportTileset"
End Sub

Sub ApplyEditedTileset(path As String)
'Apply the tileset from the tileset editor
    Call Maps(activemap).ImportTileset(path)
    
End Sub

Function hasLVLaTileset(path As String)
'Checks if the given lvl file contains a tileset
    Dim b(2) As Byte
    Dim f As Integer
    f = FreeFile
    If Dir$(path) = "" Then
        hasLVLaTileset = False
        Exit Function
    End If

    Open path For Binary As #f
    Get #f, , b
    Close #f

    If Chr(b(0)) & Chr(b(1)) = "BM" Then
        hasLVLaTileset = True
    Else
        hasLVLaTileset = False
    End If
End Function

Sub ExportTileset()
'Exports the tileset

    On Error GoTo ExportTileset_Error

    On Error GoTo errh

    'shows the common dialog
    cd.DialogTitle = "Save Tileset as bmp"
    cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    cd.Filter = "*.bmp|*.bmp"
    cd.ShowSave

    'export the tileset from the map
    Call Maps(activemap).ExportTileset(cd.filename)

    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
        MessageBox Err & " " & Err.description, vbCritical
    End If

    On Error GoTo 0
    Exit Sub

ExportTileset_Error:
    HandleError Err, "frmGeneral.ExportTileset"
End Sub

Private Sub LoadRecent()
'Load the recent entry's
'load all the menu items under window, but hide them
'until we load a map
    Dim i As Integer
    On Error GoTo LoadRecent_Error

    For i = 1 To 9
        Load mnuMaps(i)
        mnuMaps(i).visible = False



        'load the recent menu from the recent.ini if it exists
        Load mnulstRecent(i)
        mnulstRecent(i).visible = False
        If FileExists(App.path & "\recent.ini") Then
            mnulstRecent(i).Caption = INIload("Recent", CStr(i), "", App.path & "\recent.ini")

            'if there is a recent entry, make it visible
            If mnulstRecent(i).Caption <> "" Then
                mnulstRecent(i).visible = True
            End If
        End If
    Next

    'as element 0 is already preloaded, we only need to retrieve the
    'entry from the ini
    If Dir$(App.path & "\recent.ini") <> "" Then
        mnulstRecent(0).Caption = INIload("Recent", CStr(0), "", App.path & "\recent.ini")
    End If

    On Error GoTo 0
    Exit Sub

LoadRecent_Error:
    HandleError Err, "frmGeneral.LoadRecent"
End Sub

Sub AddRecent(path As String)
'Add a recent entry
    Dim i As Integer

    'if it is already in the recent, ignore it
    On Error GoTo AddRecent_Error

  
  Dim curindex As Integer
  curindex = isInRecent(path)
    If curindex <> -1 Then
      MoveRecentToTop (curindex)
    Else

    

        For i = 9 To 1 Step -1
            'move every item down
            mnulstRecent(i).Caption = mnulstRecent(i - 1).Caption
    
            'also update the visibility
            If mnulstRecent(i).Caption <> "" Then
                mnulstRecent(i).visible = True
            Else
                mnulstRecent(i).visible = False
            End If
    
            'save the entry's back into the ini
            Call INIsave("Recent", CStr(i), mnulstRecent(i).Caption, App.path & "\recent.ini")
        Next


  
  
  'now we have a free 0 element, so add the new one
    Call INIsave("Recent", CStr(0), path, App.path & "\recent.ini")
    mnulstRecent(0).Caption = path
    
      End If
      
    On Error GoTo 0
    Exit Sub

AddRecent_Error:
    HandleError Err, "frmGeneral.AddRecent"
End Sub

Sub MoveRecentToTop(Index As Integer)
'Moves a recent entry back to the top

'if it is already at the top, ignore it
    If Index = 0 Then Exit Sub

    Dim i As Integer
    Dim tmppath As String
    tmppath = mnulstRecent(Index).Caption

    For i = Index - 1 To 0 Step -1
        'move all the rest, and update the ini
        mnulstRecent(i + 1).Caption = mnulstRecent(i).Caption
        Call INIsave("Recent", CStr(i + 1), mnulstRecent(i + 1).Caption, App.path & "\recent.ini")
    Next

    'now overwrite 0 element with the given index (0 element was moved with the others)
    mnulstRecent(0).Caption = tmppath
    Call INIsave("Recent", CStr(0), tmppath, App.path & "\recent.ini")
End Sub

Function isInRecent(path As String) As Integer
'Returns the index of the file in the recent list
'Returns -1 if not found

    Dim i As Integer
    For i = 0 To 9
        If mnulstRecent(i).Caption = path Then
            'we found it, return true and exit this function
            isInRecent = i
            Exit Function
        End If
    Next

    'didn't find it
    isInRecent = -1
    Exit Function
End Function

Sub ToggleTileNr()
'Toggles the view of tile numbers on the map

    If Not loadedmaps(activemap) Then Exit Sub

    'inverse the usingtilenr
    Maps(activemap).usingtilenr = Not Maps(activemap).usingtilenr

    'update the toolbar
    Call UpdateToolBarButtons

    'also update the level
    Call Maps(activemap).UpdateLevel
End Sub

Sub ToggleGrid()
'Toggles the grid of the map

    If Not loadedmaps(activemap) Then Exit Sub

    'invert the usinggrid
    If Maps(activemap).TestMap.isRunning Then
        Maps(activemap).usinggridTest = Not Maps(activemap).usinggridTest
        Call SetSetting("ShowGridTest", CStr(CInt(Maps(activemap).usinggridTest)))
    Else
        Maps(activemap).usinggrid = Not Maps(activemap).usinggrid
        Call SetSetting("ShowGrid", CStr(CInt(Maps(activemap).usinggrid)))
    End If
    'update the toolbar
    Call UpdateToolBarButtons

    'update the preview
    Call Maps(activemap).UpdateLevel
End Sub

Sub ToggleRegions()
    If Not loadedmaps(activemap) Then Exit Sub

    Maps(activemap).ShowRegions = Not Maps(activemap).ShowRegions
    
    
    Call Maps(activemap).UpdateLevel(True, True)
    
'    Call Maps(activemap).RedrawRegions(True)
    

    'update the toolbar
    Call UpdateToolBarButtons
End Sub

Sub ToggleLVZ()
    If Not loadedmaps(activemap) Then Exit Sub

    Maps(activemap).ShowLVZ = Not Maps(activemap).ShowLVZ
        
    Call Maps(activemap).UpdateLevel(True, True)

    'update the toolbar
    Call UpdateToolBarButtons
End Sub

Friend Sub TogglePasteType(changedto As enumPasteType)
'Changes the paste type

    If Not loadedmaps(activemap) Then Exit Sub

    Maps(activemap).pastetype = changedto

    'update the selection preview on screen if there is one
    If Maps(activemap).sel.hasAlreadySelectedParts Then
        Call Maps(activemap).RedrawSelection(True)
    End If

    'update the toolbar
    Call UpdateToolBarButtons
End Sub

Sub LoadMapFromOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Load a lvl when dropped onto the form
'Use this on every OLEDragDrop for every control.
    Dim sh As Long
    Dim Pointer As Long
    Dim i As Integer
    
    Dim fileext As String
    
    On Error GoTo LoadMapFromOLEDragDrop_Error

    If Data.GetFormat(15) Then
        For i = 1 To Data.files.count
            Pointer = lOpen(Data.files(i), OF_READ)
            fileext = GetExtension(GetFileTitle(Data.files(i)))
            If fileext = "" And GetFileSize(Pointer, sh) = -1 Then
                'folder
                Exit Sub
            End If

            If fileext = "lvl" Or fileext = "elvl" Or fileext = "bak" Then
                Dim ret As Integer
                ret = OpenMap(Data.files(i))
                If ret = -1 Then
                    'no more free slots for maps
                    'stop
                    Data.files.Clear
                    Exit Sub
                End If
            End If

        Next i
        Data.files.Clear
    End If

    On Error GoTo 0
    Exit Sub

LoadMapFromOLEDragDrop_Error:
    HandleError Err, "frmGeneral.LoadMapFromOLEDragDrop"
End Sub

Sub UpdateMenuMaps()
'Update the menu Maps to the current active map
    Dim i As Integer
    
    
    For i = 0 To 9
        If loadedmaps(i) Then
            'uncheck them all
            mnuMaps(i).checked = (i = activemap)
            mnuMaps(i).Caption = Maps(i).Caption

          If Not tabloaded(i) Then
              Call tbMaps.Tabs.add(, i & "_", Maps(i).Caption)
              tabloaded(i) = True
          Else
              tbMaps.Tabs(i & "_").Caption = Maps(i).Caption
          End If
          
          If (i = activemap) Then tbMaps.Tabs(i & "_").selected = True
        Else
          If tabloaded(i) Then
              Call tbMaps.Tabs.Remove(i & "_")
              tabloaded(i) = False
          End If
        End If
    Next

    'then check the active map
'70        mnuMaps(activemap).checked = True

    
'80        Call UpdateTabMaps
End Sub
'
'Private Sub UpdateTabMaps()
'          Dim i As Integer
'
'10        tbMaps.Tabs.Clear
'20        For i = 0 To 9
'30            If loadedmaps(i) Then
'                  Dim t As MSComctlLib.Tab
'
'40                Set t = tbMaps.Tabs.add(, i & "_", Maps(i).Caption)
'
'50                If i = activemap Then
'60                    t.selected = True
'70                End If
'
'80            End If
'90        Next
'
'
'End Sub

Sub UpdateToolBarButtons()
'Update the toolbar buttons
'update action buttons depending if there is a map loaded
'    On Error GoTo UpdateToolBarButtons_Error
    On Error GoTo 0
    
    Dim i As Integer
    Dim gotSelection As Boolean
    Dim gotMap As Boolean
    
    gotMap = loadedmaps(activemap)
    
    If gotMap Then
        gotSelection = Maps(activemap).sel.hasAlreadySelectedParts
    Else
        gotSelection = False
    End If
    
    If Not gotMap Then
        
        toolbartop.Buttons("ZoomIn").Enabled = False
        toolbartop.Buttons("ZoomOut").Enabled = False
        
        
        mnuUndo.Caption = "Undo Unavailable"
        mnuRedo.Caption = "Redo Unavailable"
        
        mnuImportTileset.Enabled = False

        
'        mnuSave.Enabled = False
'        mnuSaveAs.Enabled = False
'        mnuSaveMiniMap.Enabled = False
'        mnuSaveSelect.Enabled = False
'
        mnuRevert.Enabled = False
'
'
'        mnuClose.Enabled = False
'        mnuUndo.Enabled = False
'        mnuRedo.Enabled = False
        mnuExportTileset.Enabled = False
        mnudiscardtileset.Enabled = False
'        mnuCut.Enabled = False
'        mnuCopy.Enabled = False
        mnuPaste.Enabled = False
'        mnuGrid.Enabled = False
'        mnuShowRegions.Enabled = False
'        mnuShowLVZ.Enabled = False
'        mnuTileNR.Enabled = False
'        mnuMirror.Enabled = False
'        mnuFlip.Enabled = False
'        mnuRotate.Enabled = False
'        mnuReplace.Enabled = False
'        mnuTextToMap.Enabled = False
'        mnuPTM.Enabled = False
'        mnuNormalPaste.Enabled = False
'        mnuPasteUnder.Enabled = False
'        mnuTransparentPaste.Enabled = False
'        mnuCount.Enabled = False
'        mnuNormalPaste.checked = False
'        mnuPasteUnder.checked = False
'        mnuTransparentPaste.checked = False
'        mnuWallTiles.Enabled = False
'        mnuConvtowalltiles.Enabled = False
'        mnuResize.Enabled = False
'        mnuAddToSelection.Enabled = False
'        mnuRemoveFromSelection.Enabled = False
'        mnuSelect.Enabled = False
'        mnuSelectAll.Enabled = False
'        mnuCenterSelection.Enabled = False
'        mnuCenterSelectionMap.Enabled = False
'        mnuDeleteSelection.Enabled = False
'
        mnuBookmarks.visible = False
'
'        mnuEditTileset.Enabled = False
'        mnuElvl.Enabled = False
'        mnuTileText.Enabled = False
'        mnuManageLVZ.Enabled = False
        
    Else
        
        

        mnuBookmarks.visible = True
        For i = mnuBookmark.LBound To mnuBookmark.UBound
            mnuBookmark(i).Caption = Maps(activemap).BookmarkInfo(i)
        Next

        mnuRevert.Enabled = (Maps(activemap).mapchanged And Maps(activemap).activeFile <> "")

        toolbartop.Buttons("ZoomIn").Enabled = (Maps(activemap).magnifier.zoom <> 2)
        toolbartop.Buttons("ZoomOut").Enabled = (Maps(activemap).magnifier.zoom <> 1 / 16)
        
        'update the menu items according the default tileset being used
        mnudiscardtileset.Enabled = Not Maps(activemap).usingDefaultTileset
        mnuExportTileset.Enabled = Not Maps(activemap).usingDefaultTileset
    
    
        ''''''''''''''''''''''
        'TOGGLE buttons
        
        'update according to usingtilenr
        mnuTileNR.checked = Maps(activemap).usingtilenr
        toolbartop.Buttons("TileNr").value = IIf(Maps(activemap).usingtilenr, tbrPressed, tbrUnpressed)

        'update according to usinggrid
        If Maps(activemap).TestMap.isRunning Then
            mnuGrid.checked = Maps(activemap).usinggridTest
            toolbartop.Buttons("Grid").value = IIf(Maps(activemap).usinggridTest, tbrPressed, tbrUnpressed)
        Else
            mnuGrid.checked = Maps(activemap).usinggrid
            toolbartop.Buttons("Grid").value = IIf(Maps(activemap).usinggrid, tbrPressed, tbrUnpressed)
        End If
        
        'show regions?
        mnuShowRegions.checked = Maps(activemap).ShowRegions
        toolbartop.Buttons("ShowRegions").value = IIf(Maps(activemap).ShowRegions, tbrPressed, tbrUnpressed)

        'show lvz?
        mnuShowLVZ.checked = Maps(activemap).ShowLVZ
        toolbartop.Buttons("ShowLVZ").value = IIf(Maps(activemap).ShowLVZ, tbrPressed, tbrUnpressed)
        
        

        '''''''''''
        'Draw mode
                        
        Select Case Maps(activemap).pastetype
        Case enumPasteType.p_normal
'            toolbartop.Buttons("PasteType").ButtonMenus("PasteUnder").
            toolbartop.Buttons("PasteType").Image = "PasteNormal"
            mnuNormalPaste.checked = True
            mnuTransparentPaste.checked = False
            mnuPasteUnder.checked = False
            
        Case enumPasteType.p_trans
            toolbartop.Buttons("PasteType").Image = "PasteTransparent"
            mnuTransparentPaste.checked = True
            mnuNormalPaste.checked = False
            mnuPasteUnder.checked = False
            
        Case enumPasteType.p_under
            toolbartop.Buttons("PasteType").Image = "PasteUnder"
            mnuPasteUnder.checked = True
            mnuTransparentPaste.checked = False
            mnuNormalPaste.checked = False
            
        End Select
        


        'update the redo button & menu
        mnuRedo.Enabled = (Maps(activemap).undoredo.redocurpos - 1 >= 0)

        'update the undo button & menu
        mnuUndo.Enabled = (Maps(activemap).undoredo.undocurpos - 1 >= 0)
        
        mnuUndo.Caption = "Undo " & Maps(activemap).undoredo.GetUndoComment
        mnuRedo.Caption = "Redo " & Maps(activemap).undoredo.GetRedoComment
        
        'update the paste button
        mnuPaste.Enabled = SharedVar.clipHasData


        
        'Check for a valid wall tile set to enable the convert to WT menu
        mnuConvtowalltiles.Enabled = False
        If gotSelection Then
            For i = 0 To 7
                If Maps(activemap).walltiles.isValidSet(i) Then
                    mnuConvtowalltiles.Enabled = True
                End If
            Next
        End If

        
        Call UpdateAutoSaveList
    End If
    
    '''''''''''''''''''''''''''''
    'Stuff depending only if there's a map or not

    mnuDrawMode.Enabled = gotMap
    
    mnuImportTileset.Enabled = gotMap
    mnuSave.Enabled = gotMap
    mnuSaveAs.Enabled = gotMap
    mnuSaveMiniMap.Enabled = gotMap
    mnuSaveSelect.Enabled = gotMap
    mnuClose.Enabled = gotMap
    mnuGrid.Enabled = gotMap
    mnuShowLVZ.Enabled = gotMap
    mnuShowRegions.Enabled = gotMap
    mnuTileNR.Enabled = gotMap
    mnuReplace.Enabled = gotMap
    mnuTextToMap.Enabled = gotMap
    mnuPTM.Enabled = gotMap
    mnuNormalPaste.Enabled = gotMap
    mnuPasteUnder.Enabled = gotMap
    mnuTransparentPaste.Enabled = gotMap
    mnuCount.Enabled = gotMap
    mnuWallTiles.Enabled = gotMap
    mnuSelect.Enabled = gotMap
    mnuSelectAll.Enabled = gotMap

    mnuEditTileset.Enabled = gotMap
    mnuElvl.Enabled = gotMap
    mnuTileText.Enabled = gotMap
    mnuManageLVZ.Enabled = gotMap
    
    '''''''''''''''''''''''''''''
    'Stuff depending on selection

    mnuCut.Enabled = gotSelection
    mnuCopy.Enabled = gotSelection
    mnuFlip.Enabled = gotSelection
    mnuMirror.Enabled = gotSelection
    mnuRotate.Enabled = gotSelection
    mnuResize.Enabled = gotSelection
    
    mnuSelectNone.Enabled = gotSelection
    mnuCenterSelection.Enabled = gotSelection
    mnuCenterSelectionMap.Enabled = gotSelection
    mnuAddToSelection.Enabled = gotSelection
    mnuRemoveFromSelection.Enabled = gotSelection
    mnuDeleteSelection.Enabled = gotSelection
    
    'Check for a valid wall tile set to enable the convert to WT menu
    mnuConvtowalltiles.Enabled = False
    If gotSelection Then
        For i = 0 To 7
            If Maps(activemap).walltiles.isValidSet(i) Then
                mnuConvtowalltiles.Enabled = True
            End If
        Next
    End If
        
    
    
    
    '''''''''''''''''''''''''''''''''
    'Things that depend on menu items
    
    tlbTileset.Buttons("EditTileset").Enabled = mnuEditTileset.Enabled
    tlbTileset.Buttons("ImportTileset").Enabled = mnuImportTileset.Enabled
    tlbTileset.Buttons("ExportTileset").Enabled = mnuExportTileset.Enabled
    tlbTileset.Buttons("DiscardTileset").Enabled = mnudiscardtileset.Enabled
    tlbTileset.Buttons("EditWalltiles").Enabled = mnuWallTiles.Enabled
        
    toolbartop.Buttons("Cut").Enabled = mnuCut.Enabled
    toolbartop.Buttons("Copy").Enabled = mnuCopy.Enabled

    toolbartop.Buttons("Flip").Enabled = mnuFlip.Enabled
    toolbartop.Buttons("Mirror").Enabled = mnuMirror.Enabled
    toolbartop.Buttons("Rotate").Enabled = mnuRotate.Enabled
    

    toolbartop.Buttons("Redo").Enabled = mnuRedo.Enabled
    toolbartop.Buttons("Undo").Enabled = mnuUndo.Enabled
    toolbartop.Buttons("Undo").tooltiptext = mnuUndo.Caption
    toolbartop.Buttons("Redo").tooltiptext = mnuRedo.Caption
    
    toolbartop.Buttons("Save").Enabled = mnuSave.Enabled
    
    
    toolbartop.Buttons("Paste").Enabled = mnuPaste.Enabled
    toolbartop.Buttons("Grid").Enabled = mnuGrid.Enabled
    toolbartop.Buttons("TileNr").Enabled = mnuTileNR.Enabled
    toolbartop.Buttons("Replace").Enabled = mnuReplace.Enabled
    toolbartop.Buttons("TextToMap").Enabled = mnuTextToMap.Enabled
    toolbartop.Buttons("PicToMap").Enabled = mnuPTM.Enabled
    toolbartop.Buttons("PasteType").Enabled = mnuDrawMode.Enabled
    toolbartop.Buttons("EditELVL").Enabled = mnuElvl.Enabled
    tlbTileset.Buttons("EditLVZ").Enabled = mnuManageLVZ.Enabled

    toolbartop.Buttons("ShowRegions").Enabled = mnuShowRegions.Enabled
    toolbartop.Buttons("ShowLVZ").Enabled = mnuShowLVZ.Enabled
    
        
    '''''''''''''''
    'Toolbars stuff
    
    If CBool(GetSetting("ShowToolbarStandard", "1")) Then
        mnuToolbarStandard.checked = True
        toolbartop.visible = True
    Else
        mnuToolbarStandard.checked = False
        toolbartop.visible = False
    End If

    If CBool(GetSetting("ShowToolbarTools", "1")) Then
        mnuToolbarTools.checked = True
        Toolbarleft.visible = True
    Else
        mnuToolbarTools.checked = False
        Toolbarleft.visible = False
    End If

    If CBool(GetSetting("ShowToolbarMapTabs", "1")) Then
        mnuToolbarMapTabs.checked = True
        tlbTabs.visible = True
    Else
        mnuToolbarMapTabs.checked = False
        tlbTabs.visible = False
    End If
    
    tlbTabs.Align = CInt(GetSetting("MapTabPosition", tlbTabs.Align))
    If tlbTabs.Align = vbAlignBottom Then
        tbMaps.Placement = tabPlacementBottom
        cmdChangeTabPos.Caption = "Ù"
    Else
        tlbTabs.Align = vbAlignTop
        tbMaps.Placement = tabPlacementTop
        cmdChangeTabPos.Caption = "Ú"
    End If
    
    
    If CBool(GetSetting("PinToolOptions", "1")) Then
        mnuTogglePinToolOptions.checked = True
    Else
        mnuTogglePinToolOptions.checked = False
    End If

    If CBool(GetSetting("PinToolOptions", "1")) Or CBool(GetSetting("ShowToolbarToolOptions", "1")) And toolOptionExist(curtool) Then
        mnuToolbarToolOptions.checked = True
        tlbToolOptions.visible = True

        If toolOptionExist(curtool) Then
            frmTool(curtool - 1).visible = True
        End If
    Else
        mnuToolbarToolOptions.checked = False
        tlbToolOptions.visible = False
        If toolOptionExist(curtool) Then
            frmTool(curtool - 1).visible = False
        End If
    End If

    
'    Exit Sub
'
'UpdateToolBarButtons_Error:
'    HandleError Err, "frmGeneral.UpdateToolBarButtons"
End Sub

Sub UpdateToolToolbar()
'Update the info on the tools toolbar
'          Dim i As Integer

    Select Case curtool
    Case T_magnifier    '0
        If loadedmaps(activemap) Then
            Dim zoom1 As Integer
            Dim zoom2 As Integer
            If Maps(activemap).currentzoom > 1 Then
                zoom1 = Maps(activemap).currentzoom
            Else
                zoom1 = 1
            End If

            If Maps(activemap).currentzoom < 1 Then
                zoom2 = 1 / (Maps(activemap).currentzoom)
            Else
                zoom2 = 1
            End If

            lblcurrentzoom.Caption = "Zoom: " & zoom1 & ":" & zoom2
        Else
            lblcurrentzoom.Caption = ""
        End If
    Case T_magicwand    '2
        chkMagicWandDiagonal.value = CInt(GetSetting("WandDiagonal", "0"))
        chkMagicWandScreen.value = CInt(GetSetting("WandScreen", "1"))

    Case T_airbrush
        chkUseAsAsteroidBrush.value = CInt(GetSetting("UseAirBrushAsAsteroids", "0"))
        chkuseSmallAsteroids1.value = CInt(GetSetting("UseSmallAsteroids1", "1"))
        chkUseSmallAsteroids2.value = CInt(GetSetting("UseSmallAsteroids2", "1"))
        chkUseBigAsteroids.value = CInt(GetSetting("UseBigAsteroids", "1"))
        chkuseSmallAsteroids1.visible = CBool(chkUseAsAsteroidBrush.value)
        chkUseSmallAsteroids2.visible = CBool(chkUseAsAsteroidBrush.value)
        chkUseBigAsteroids.visible = CBool(chkUseAsAsteroidBrush.value)
        If chkUseAsAsteroidBrush.value = vbChecked Then
            frmAsteroids.width = 4815
        Else
            frmAsteroids.width = 2175
        End If
        frmDensity.Left = frmAsteroids.Left + frmAsteroids.width + 125
        frmSize.Left = frmDensity.Left + frmDensity.width + 125

        sldAirbDensity.value = CInt(GetSetting("AirbDensity", "100"))
        sldAirbSize.value = CInt(GetSetting("AirbSize", "10"))
        txtAirBrDensity.Text = sldAirbDensity.value
        txtAirBrSize.Text = sldAirbSize.value

    Case T_bucket
        chkFillDiagonal.value = CInt(GetSetting("FillDiagonal", "0"))
        chkFillInScreen.value = CInt(GetSetting("FillInScreen", "1"))

    Case T_pencil
        toolSize(curtool - 1).value = CInt(GetSetting("ToolSize" & ToolName(CInt(curtool - 1)), "1"))

        If CInt(GetSetting("ToolTip" & ToolName(CInt(curtool - 1)), "1")) = 1 Then
            optToolSquare(curtool - 1).value = True
        Else
            optToolRound(curtool - 1).value = True
        End If
        
        chkAdvancedPencil.value = CInt(GetSetting("AdvancedPencil", "0"))

    Case T_Eraser, T_replacebrush
        toolSize(curtool - 1).value = CInt(GetSetting("ToolSize" & ToolName(CInt(curtool - 1)), "1"))

        If CInt(GetSetting("ToolTip" & ToolName(CInt(curtool - 1)), "1")) = 1 Then
            optToolSquare(curtool - 1).value = True
        Else
            optToolRound(curtool - 1).value = True
        End If

    Case T_line, T_rectangle, T_filledrectangle, t_spline
        toolSize(curtool - 1).value = CInt(GetSetting("ToolSize" & ToolName(CInt(curtool - 1)), "1"))

        If CInt(GetSetting("ToolTip" & ToolName(CInt(curtool - 1)), "1")) = 1 Then
            optToolSquare(curtool - 1).value = True
        Else
            optToolRound(curtool - 1).value = True
        End If

        toolStep(curtool - 1).value = CInt(GetSetting("ToolStep" & ToolName(CInt(curtool - 1)), "0"))

    Case T_ellipse, T_filledellipse
        toolSize(curtool - 1).value = CInt(GetSetting("ToolSize" & ToolName(CInt(curtool - 1)), "1"))
        toolStep(curtool - 1).value = CInt(GetSetting("ToolStep" & ToolName(CInt(curtool - 1)), "0"))
        chkRenderAfter(curtool - 1).value = CInt(GetSetting("ToolRenderAfter" & ToolName(CInt(curtool - 1)), "0"))

    Case T_customshape
        chkRenderAfter(curtool - 1).value = CInt(GetSetting("ToolRenderAfter" & ToolName(CInt(curtool - 1)), "0"))
        Select Case curCustomShape
        Case s_cogwheel, s_star
'            customShapeAngle(curCustomShape).value = CInt(GetSetting("CustomShapeAngle" & CustomShapeName(curCustomShape), "0"))
            customShapeSize(curCustomShape).value = CInt(GetSetting("CustomShapeSize" & CustomShapeName(curCustomShape), "1"))
            customShapeTeethNumber(curCustomShape).value = CInt(GetSetting("CustomShapeTeethNumber" & CustomShapeName(curCustomShape), "5"))
            customShapeTeethSize(curCustomShape).value = CInt(GetSetting("CustomShapeTeethSize" & CustomShapeName(curCustomShape), "50"))
        Case s_regular
'            customShapeAngle(curCustomShape).value = CInt(GetSetting("CustomShapeAngle" & CustomShapeName(curCustomShape), "0"))
            customShapeSize(curCustomShape).value = CInt(GetSetting("CustomShapeSize" & CustomShapeName(curCustomShape), "1"))
            customShapeTeethNumber(curCustomShape).value = CInt(GetSetting("CustomShapeTeethNumber" & CustomShapeName(curCustomShape), "5"))
        End Select
        Call UpdateCustomShapePreview(curCustomShape)

    Case T_Region

        optRegionSel(0).value = CBool(GetSetting("RegionUseWand", "0"))
        optRegionSel(1).value = Not optRegionSel(0).value

        Call UpdateRegionList

    Case T_TestMap
        chkTileCollision.value = CInt(GetSetting("TileCollision", "1"))
        optShip(CInt(GetSetting("TestMapShip", "1")) - 1).value = True
  
  Case T_dropper
  
      chkDropperIgnoreEmpty.value = CInt(GetSetting("DropEmpty", 1))
      
    Case T_lvz
      If loadedmaps(activemap) Then txtLvzSnap.Text = Maps(activemap).lvz.snap
      
'        Dim oldIdx As Integer
'        oldIdx = cmbMapObjects.ListIndex
        
'        Dim MapObjCount As Integer
'        MapObjCount = 0
'        chkLvzSnapToTiles.Value = CInt(GetSetting("SnapToTiles", "1"))
'900           chkLvzForceTransparency.value = CInt(GetSetting("ForceLVZTransparency", "1"))
'        cmbMapObjects.Clear
'        Dim j As Integer
        
'        If loadedmaps(activemap) Then
'            For i = 0 To Maps(activemap).lvz.getLVZCount - 1
'                cmbMapObjects.addItem Maps(activemap).lvz.getLVZ(i).name
'                For j = 0 To Maps(activemap).lvz.getLVZ(i).mapObjectCount - 1
'                    cmbMapObjects.addItem "    Mapobject" & j
'                    MapObjCount = MapObjCount + 1
'                Next
'            Next
'        End If
        
'        If MapObjCount = 0 Then
'            cmbMapObjects.Enabled = False
'            cmdLVZJumpTo.Enabled = False
'            cmdLVZMoveToScreen.Enabled = False
'            chkLvzSnapToTiles.Enabled = False
'            chkLvzForceTransparency.Enabled = False
'        Else
'            cmbMapObjects.Enabled = True
'            cmdLVZJumpTo.Enabled = True
'            cmdLVZMoveToScreen.Enabled = True
'            chkLvzSnapToTiles.Enabled = True
'            chkLvzForceTransparency.Enabled = True
'
'
'            If oldIdx = -1 Then
'                oldIdx = 0
'            End If
'            'look for first map object starting from oldix
'            For i = oldIdx To cmbMapObjects.ListCount - 1
'                If Mid$(cmbMapObjects.list(i), 1, 4) = "    " Then
'                    cmbMapObjects.ListIndex = i
'                    Exit Sub
'                End If
'            Next
'
'            'no mapobject found. There is one, else we would have been disabled.
'            'start looking backwards
'            For i = oldIdx To 0 Step -1
'                If Mid$(cmbMapObjects.list(i), 1, 4) = "    " Then
'                    cmbMapObjects.ListIndex = i
'                    Exit Sub
'                End If
'            Next
'        End If
        
    End Select

    If loadedmaps(activemap) Then
        Call Maps(activemap).UpdateScrollbars(True)
    End If
End Sub


'Private Function checkToolOptionExist(ByVal tool As toolenum) As Boolean
'    On Error GoTo errh
'    Dim tmp As String
'
'    tmp = frmTool(tool - 1).Caption
'
'    checkToolOptionExist = True
'    Exit Function
'errh:
'    If Err = 340 Then
'        checkToolOptionExist = False
'        Exit Function
'    End If
'End Function

Function toolOptionExist(tool As toolenum) As Boolean
'Checks if the tool option frame for the toolbar exists
    toolOptionExist = c_toolOptionExists(tool)
End Function

Function customShapeOptionExist(tool As customshapeEnum) As Boolean
'Checks if the tool option frame for the toolbar exists
    On Error GoTo errh
    Dim tmp As String

    tmp = frmCustomShape(tool).Caption

    customShapeOptionExist = True
    Exit Function
errh:
    If Err = 340 Then
        customShapeOptionExist = False
        Exit Function
    End If
End Function




Sub UpdatePreview()
    Dim Button As Integer
    
    For Button = 1 To 2
        
        'Draw the large preview
        Call cTileset.DrawTilePreview(Button, pictilesetlarge, shppreviewsel(Button).Left + 1, shppreviewsel(Button).Top + 1, shppreviewsel(Button).width - 1, shppreviewsel(Button).height - 2)
        
        'Draw the small preview
        Call cTileset.DrawTilePreview(Button, picsmalltilepreview, shpsmalltilepreview(Button).Left + 1, shpsmalltilepreview(Button).Top + 1, shpsmalltilepreview(Button).width - 1, shpsmalltilepreview(Button).height - 2)
        
    Next
'                SetStretchBltMode pictilesetlarge.hDC, HALFTONE
'                StretchBlt pictilesetlarge.hDC, shppreviewsel(button).Left + 1, shppreviewsel(button).Top + 1, shppreviewsel(button).Width - 2, shppreviewsel(button).Height - 2, picWalltiles.hDC, (.tileset.selection(button).group Mod 4) * 64, (.tileset.selection(button).group \ 4) * 64, 4 * TILEW, 4 * TILEW, vbSrcCopy
'
'                'update the smalltile preview, used when the right panel is shrinked
'                BitBlt picsmalltilepreview.hDC, 1, 1 + (button - 1) * 18, TILEW, TILEW, pictileset.hDC, shptilesel(button).Left, shptilesel(button).Top, vbSrcCopy
'
'            ElseIf .tileset.selection(button).selectionType = TS_Tiles Then
'                SetStretchBltMode pictilesetlarge.hDC, COLORONCOLOR
'                StretchBlt pictilesetlarge.hDC, shppreviewsel(button).Left + 1, shppreviewsel(button).Top + 1, shppreviewsel(button).Width - 2, shppreviewsel(button).Height - 2, pictileset.hDC, shptilesel(button).Left, shptilesel(button).Top, TILEW, TILEW, vbSrcCopy
'
'                'update the smalltile preview, used when the right panel is shrinked
'                BitBlt picsmalltilepreview.hDC, 1, 1 + (button - 1) * 18, TILEW, TILEW, pictileset.hDC, shptilesel(button).Left, shptilesel(button).Top, vbSrcCopy
'
'            ElseIf .tileset.selection(button).selectionType = TS_LVZ Then
'
'                Dim lvzidx As Integer
'                Dim imgidx As Integer
'                Dim srcDC As Long
'
'                lvzidx = .tileset.selection(button).group
'                imgidx = .tileset.selection(button).tilenr
'                srcDC = .lvz.pichDClib(.lvz.getImageDefinition(lvzidx, imgidx).picboxIdx)
'                SetStretchBltMode pictilesetlarge.hDC, HALFTONE
'                StretchBlt pictilesetlarge.hDC, shppreviewsel(button).Left + 1, shppreviewsel(button).Top + 1, shppreviewsel(button).Width - 2, shppreviewsel(button).Height - 2, srcDC, 0, 0, .tileset.selection(button).pixelSize.X, .tileset.selection(button).pixelSize.Y, vbSrcCopy
'
'                StretchBlt picsmalltilepreview.hDC, 1, 1 + (button - 1) * 18, TILEW, TILEW, srcDC, 0, 0, .tileset.selection(button).pixelSize.X, .tileset.selection(button).pixelSize.Y, vbSrcCopy
'
'
End Sub

Private Sub SetIconsToMenus()
'Set all icons in the menus
'first add a space to each menu item so there is space
'between the icon and the text
    Dim c As Control
'    On Error GoTo SetIconsToMenus_Error

    For Each c In frmGeneral.Controls
        If TypeOf c Is Menu Then
            If c.Caption <> "-" Then
                c.Caption = " " & c.Caption
            End If
        End If
    Next

    optRegionSel(0).Picture = imlToolbarIcons.ListImages("MagicWand").Picture
    optRegionSel(1).Picture = imlToolbarIcons.ListImages("Selection").Picture

    'new icon
    Dim mainmnu As Long
    Dim hmnu(6) As Long
    mainmnu = GetMenu(Me.hWnd)

    Dim i As Integer
    For i = 0 To UBound(hmnu)
        hmnu(i) = GetSubMenu(mainmnu, i)
    Next


    '---------------------- FILE -----------
    'New
    picicons(0).Picture = imlToolbarIcons.ListImages("New").Picture
    picicons(0).Picture = picicons(0).Image
    SetMenuItemBitmaps hmnu(0), 0, MF_BYPOSITION, picicons(0).Picture, picicons(0).Picture

    'Open
    Load picicons(1)
    picicons(1).Picture = imlToolbarIcons.ListImages("Open").Picture
    picicons(1).Picture = picicons(1).Image
    SetMenuItemBitmaps hmnu(0), 1, MF_BYPOSITION, picicons(1).Picture, picicons(1).Picture

    'Save
    Load picicons(2)
    picicons(2).Picture = imlToolbarIcons.ListImages("Save").Picture
    picicons(2).Picture = picicons(2).Image
    SetMenuItemBitmaps hmnu(0), 3, MF_BYPOSITION, picicons(2).Picture, picicons(2).Picture

    'import tileset
    Load picicons(38)
    picicons(38).Picture = imlToolbarIcons.ListImages("ImportTileset").Picture
    picicons(38).Picture = picicons(38).Image
    SetMenuItemBitmaps hmnu(0), 9, MF_BYPOSITION, picicons(38).Picture, picicons(38).Picture

    'export tileset
    Load picicons(39)
    picicons(39).Picture = imlToolbarIcons.ListImages("ExportTileset").Picture
    picicons(39).Picture = picicons(39).Image
    SetMenuItemBitmaps hmnu(0), 10, MF_BYPOSITION, picicons(39).Picture, picicons(39).Picture

    'discard tileset
    Load picicons(40)
    picicons(40).Picture = imlToolbarIcons.ListImages("DiscardTileset").Picture
    picicons(40).Picture = picicons(40).Image
    SetMenuItemBitmaps hmnu(0), 11, MF_BYPOSITION, picicons(40).Picture, picicons(40).Picture

    '----------------------------------------------

    '-------------------- EDIT ------------------
    'Undo
    Load picicons(3)
    picicons(3).Picture = imlToolbarIcons.ListImages("Undo").Picture
    picicons(3).Picture = picicons(3).Image
    SetMenuItemBitmaps hmnu(1), 0, MF_BYPOSITION, picicons(3).Picture, picicons(3).Picture

    'Redo
    Load picicons(4)
    picicons(4).Picture = imlToolbarIcons.ListImages("Redo").Picture
    picicons(4).Picture = picicons(4).Image
    SetMenuItemBitmaps hmnu(1), 1, MF_BYPOSITION, picicons(4).Picture, picicons(4).Picture

    'Cut
    Load picicons(5)
    picicons(5).Picture = imlToolbarIcons.ListImages("Cut").Picture
    picicons(5).Picture = picicons(5).Image
    SetMenuItemBitmaps hmnu(1), 3, MF_BYPOSITION, picicons(5).Picture, picicons(5).Picture

    'Copy
    Load picicons(6)
    picicons(6).Picture = imlToolbarIcons.ListImages("Copy").Picture
    picicons(6).Picture = picicons(6).Image
    SetMenuItemBitmaps hmnu(1), 4, MF_BYPOSITION, picicons(6).Picture, picicons(6).Picture

    'Paste
    Load picicons(7)
    picicons(7).Picture = imlToolbarIcons.ListImages("Paste").Picture
    picicons(7).Picture = picicons(7).Image
    SetMenuItemBitmaps hmnu(1), 5, MF_BYPOSITION, picicons(7).Picture, picicons(7).Picture

    'Switch/Replace
    Load picicons(8)
    picicons(8).Picture = imlToolbarIcons.ListImages("Replace").Picture
    picicons(8).Picture = picicons(8).Image
    SetMenuItemBitmaps hmnu(1), 7, MF_BYPOSITION, picicons(8).Picture, picicons(8).Picture

    'FlipH
    Load picicons(9)
    picicons(9).Picture = imlToolbarIcons.ListImages("Mirror").Picture
    picicons(9).Picture = picicons(9).Image
    SetMenuItemBitmaps hmnu(1), 8, MF_BYPOSITION, picicons(9).Picture, picicons(9).Picture

    'FlipV
    Load picicons(10)
    picicons(10).Picture = imlToolbarIcons.ListImages("Flip").Picture
    picicons(10).Picture = picicons(10).Image
    SetMenuItemBitmaps hmnu(1), 9, MF_BYPOSITION, picicons(10).Picture, picicons(10).Picture

    'Rotate
    Load picicons(11)
    picicons(11).Picture = imlToolbarIcons.ListImages("Rotate").Picture
    picicons(11).Picture = picicons(11).Image
    SetMenuItemBitmaps hmnu(1), 10, MF_BYPOSITION, picicons(11).Picture, picicons(11).Picture

    'Resize
    Load picicons(44)
    picicons(44).Picture = imlToolbarIcons.ListImages("Resize").Picture
    picicons(44).Picture = picicons(44).Image
    SetMenuItemBitmaps hmnu(1), 11, MF_BYPOSITION, picicons(44).Picture, picicons(44).Picture

    'walltiles
    Load picicons(34)
    picicons(34).Picture = imlToolbarIcons.ListImages("WallTiles").Picture
    picicons(34).Picture = picicons(34).Image
    Load picicons(35)
    picicons(35).Picture = imlToolbarIcons.ListImages("WallTilesSelected").Picture
    picicons(35).Picture = picicons(35).Image

    SetMenuItemBitmaps hmnu(1), 13, MF_BYPOSITION, picicons(34).Picture, picicons(35).Picture

    'edit tileset
    Load picicons(54)
    picicons(54).Picture = imlToolbarIcons.ListImages("EditTileset").Picture
    picicons(54).Picture = picicons(54).Image
    SetMenuItemBitmaps hmnu(1), 14, MF_BYPOSITION, picicons(54).Picture, picicons(54).Picture

    'edit tiletext
    Load picicons(57)
    picicons(57).Picture = imlToolbarIcons.ListImages("EditTileText").Picture
    picicons(57).Picture = picicons(57).Image
    SetMenuItemBitmaps hmnu(1), 15, MF_BYPOSITION, picicons(57).Picture, picicons(57).Picture

    'Count
    Load picicons(29)
    picicons(29).Picture = imlToolbarIcons.ListImages("Count").Picture
    picicons(29).Picture = picicons(29).Image

    SetMenuItemBitmaps hmnu(1), 17, MF_BYPOSITION, picicons(29).Picture, picicons(29).Picture

    'TTM
    Load picicons(12)
    picicons(12).BackColor = vbWhite
    picicons(12).Picture = imlToolbarIcons.ListImages("TTM").Picture
    picicons(12).Picture = picicons(12).Image
    SetMenuItemBitmaps hmnu(1), 18, MF_BYPOSITION, picicons(12).Picture, picicons(12).Picture

    'PTM
    Load picicons(13)
    picicons(13).Picture = imlToolbarIcons.ListImages("PTM").Picture
    picicons(13).Picture = picicons(13).Image
    SetMenuItemBitmaps hmnu(1), 19, MF_BYPOSITION, picicons(13).Picture, picicons(13).Picture

    'eLVL
    Load picicons(56)
    picicons(56).Picture = imlToolbarIcons.ListImages("elvl").Picture
    picicons(56).Picture = picicons(56).Image
    SetMenuItemBitmaps hmnu(1), 21, MF_BYPOSITION, picicons(56).Picture, picicons(56).Picture

    'LVZ
    Load picicons(62)
    picicons(62).Picture = imlToolbarIcons.ListImages("Package").Picture
    picicons(62).Picture = picicons(62).Image
    SetMenuItemBitmaps hmnu(1), 22, MF_BYPOSITION, picicons(62).Picture, picicons(62).Picture

    '-----------------------------------------

    ' --------------- OPTIONS -------------------
    'Grid
    Load picicons(14)
    picicons(14).Picture = imlToolbarIcons.ListImages("Grid").Picture
    picicons(14).Picture = picicons(14).Image
    Load picicons(16)
    picicons(16).Picture = imlToolbarIcons.ListImages("GridSelected").Picture
    picicons(16).Picture = picicons(16).Image

    SetMenuItemBitmaps hmnu(2), 0, MF_BYPOSITION, picicons(14).Picture, picicons(16).Picture

    'Tilenr
    Load picicons(15)
    picicons(15).Picture = imlToolbarIcons.ListImages("Tilenr").Picture
    picicons(15).Picture = picicons(15).Image
    Load picicons(17)
    picicons(17).Picture = imlToolbarIcons.ListImages("TilenrSelected").Picture
    picicons(17).Picture = picicons(17).Image

    SetMenuItemBitmaps hmnu(2), 1, MF_BYPOSITION, picicons(15).Picture, picicons(17).Picture

    'regions
    Load picicons(58)
    picicons(58).Picture = imlToolbarIcons.ListImages("Regions").Picture
    picicons(58).Picture = picicons(58).Image
    Load picicons(59)
    picicons(59).Picture = imlToolbarIcons.ListImages("RegionsSelected").Picture
    picicons(59).Picture = picicons(59).Image

    SetMenuItemBitmaps hmnu(2), 2, MF_BYPOSITION, picicons(58).Picture, picicons(59).Picture

    'lvz
    Load picicons(60)
    picicons(60).Picture = imlToolbarIcons.ListImages("Package").Picture
    picicons(60).Picture = picicons(60).Image
    Load picicons(61)
    picicons(61).Picture = imlToolbarIcons.ListImages("PackageSelected").Picture
    picicons(61).Picture = picicons(61).Image

    SetMenuItemBitmaps hmnu(2), 3, MF_BYPOSITION, picicons(60).Picture, picicons(61).Picture

'    'Normal paste
'    Load picicons(22)
'    picicons(22).Picture = imlToolbarIcons.ListImages("PasteNormal").Picture
'    picicons(22).Picture = picicons(22).Image
'    Load picicons(23)
'    picicons(23).Picture = imlToolbarIcons.ListImages("Paste NormalSelected").Picture
'    picicons(23).Picture = picicons(23).Image
'
'    SetMenuItemBitmaps hmnu(2), 5, MF_BYPOSITION, picicons(22).Picture, picicons(23).Picture
'
'    'Paste Under
'    Load picicons(24)
'    picicons(24).Picture = imlToolbarIcons.ListImages("PasteUnder").Picture
'    picicons(24).Picture = picicons(24).Image
'    Load picicons(25)
'    picicons(25).Picture = imlToolbarIcons.ListImages("Paste UnderSelected").Picture
'    picicons(25).Picture = picicons(25).Image
'
'    SetMenuItemBitmaps hmnu(2), 6, MF_BYPOSITION, picicons(24).Picture, picicons(25).Picture
'
'    'Paste Transparent
'    Load picicons(26)
'    picicons(26).Picture = imlToolbarIcons.ListImages("PasteTransparent").Picture
'    picicons(26).Picture = picicons(26).Image
'    Load picicons(27)
'    picicons(27).Picture = imlToolbarIcons.ListImages("Paste TransparentSelected").Picture
'    picicons(27).Picture = picicons(27).Image
'
'    SetMenuItemBitmaps hmnu(2), 7, MF_BYPOSITION, picicons(26).Picture, picicons(27).Picture

    'Options
    Load picicons(33)
    picicons(33).Picture = imlToolbarIcons.ListImages("Options").Picture
    picicons(33).Picture = picicons(33).Image

    SetMenuItemBitmaps hmnu(2), 7, MF_BYPOSITION, picicons(33).Picture, picicons(33).Picture
    '---------------------------------------

    '------------------ SELECTION ----------
    Dim selectsub As Long
    selectsub = GetSubMenu(hmnu(3), 0)

    Dim addsub As Long
    addsub = GetSubMenu(hmnu(3), 1)

    Dim removesub As Long
    removesub = GetSubMenu(hmnu(3), 2)

    'Select All
    Load picicons(46)
    picicons(46).Picture = imlToolbarIcons.ListImages("SelectAll").Picture
    picicons(46).Picture = picicons(46).Image

    SetMenuItemBitmaps selectsub, 1, MF_BYPOSITION, picicons(46).Picture, picicons(46).Picture
    SetMenuItemBitmaps addsub, 0, MF_BYPOSITION, picicons(46).Picture, picicons(46).Picture

    'Select None
    Load picicons(47)
    picicons(47).Picture = imlToolbarIcons.ListImages("SelectNone").Picture
    picicons(47).Picture = picicons(47).Image

    SetMenuItemBitmaps removesub, 0, MF_BYPOSITION, picicons(47).Picture, picicons(47).Picture
    SetMenuItemBitmaps removesub, 1, MF_BYPOSITION, picicons(47).Picture, picicons(47).Picture

    'Select
    Load picicons(41)
    picicons(41).Picture = imlToolbarIcons.ListImages("Select").Picture
    picicons(41).Picture = picicons(41).Image

    SetMenuItemBitmaps hmnu(3), 0, MF_BYPOSITION, picicons(41).Picture, picicons(41).Picture

    'Select Left Tile
    Load picicons(52)
    picicons(52).Picture = imlToolbarIcons.ListImages("SelectLeftTile").Picture
    picicons(52).Picture = picicons(52).Image

    SetMenuItemBitmaps selectsub, 3, MF_BYPOSITION, picicons(52).Picture, picicons(52).Picture
    SetMenuItemBitmaps selectsub, 4, MF_BYPOSITION, picicons(52).Picture, picicons(52).Picture

    'Select Right Tile
    Load picicons(53)
    picicons(53).Picture = imlToolbarIcons.ListImages("SelectRightTile").Picture
    picicons(53).Picture = picicons(53).Image

    SetMenuItemBitmaps selectsub, 5, MF_BYPOSITION, picicons(53).Picture, picicons(53).Picture
    SetMenuItemBitmaps selectsub, 6, MF_BYPOSITION, picicons(53).Picture, picicons(53).Picture

    'Select Add
    Load picicons(42)
    picicons(42).Picture = imlToolbarIcons.ListImages("SelectAdd").Picture
    picicons(42).Picture = picicons(42).Image

    SetMenuItemBitmaps hmnu(3), 1, MF_BYPOSITION, picicons(42).Picture, picicons(42).Picture

    'Select Add Left Tile
    Load picicons(48)
    picicons(48).Picture = imlToolbarIcons.ListImages("SelectAddLeftTile").Picture
    picicons(48).Picture = picicons(48).Image

    SetMenuItemBitmaps addsub, 2, MF_BYPOSITION, picicons(48).Picture, picicons(48).Picture
    SetMenuItemBitmaps addsub, 3, MF_BYPOSITION, picicons(48).Picture, picicons(48).Picture

    'Select Add Right Tile
    Load picicons(49)
    picicons(49).Picture = imlToolbarIcons.ListImages("SelectAddRightTile").Picture
    picicons(49).Picture = picicons(49).Image

    SetMenuItemBitmaps addsub, 4, MF_BYPOSITION, picicons(49).Picture, picicons(49).Picture
    SetMenuItemBitmaps addsub, 5, MF_BYPOSITION, picicons(49).Picture, picicons(49).Picture

    'Select Remove
    Load picicons(43)
    picicons(43).Picture = imlToolbarIcons.ListImages("SelectRemove").Picture
    picicons(43).Picture = picicons(43).Image

    SetMenuItemBitmaps hmnu(3), 2, MF_BYPOSITION, picicons(43).Picture, picicons(43).Picture

    'Select Remove Left Tile
    Load picicons(50)
    picicons(50).Picture = imlToolbarIcons.ListImages("SelectRemoveLeftTile").Picture
    picicons(50).Picture = picicons(50).Image

    SetMenuItemBitmaps removesub, 3, MF_BYPOSITION, picicons(50).Picture, picicons(50).Picture
    SetMenuItemBitmaps removesub, 4, MF_BYPOSITION, picicons(50).Picture, picicons(50).Picture

    'Select remove Right Tile
    Load picicons(51)
    picicons(51).Picture = imlToolbarIcons.ListImages("SelectRemoveRightTile").Picture
    picicons(51).Picture = picicons(51).Image

    SetMenuItemBitmaps removesub, 5, MF_BYPOSITION, picicons(51).Picture, picicons(51).Picture
    SetMenuItemBitmaps removesub, 6, MF_BYPOSITION, picicons(51).Picture, picicons(51).Picture

    'Convert to walltiles
    Load picicons(36)
    picicons(36).Picture = imlToolbarIcons.ListImages("ConvertToWalltiles").Picture
    picicons(36).Picture = picicons(36).Image

    SetMenuItemBitmaps hmnu(3), 4, MF_BYPOSITION, picicons(36).Picture, picicons(36).Picture

    'Center in screen
    Load picicons(45)
    picicons(45).Picture = imlToolbarIcons.ListImages("CenterInScreen").Picture
    picicons(45).Picture = picicons(45).Image

    SetMenuItemBitmaps hmnu(3), 7, MF_BYPOSITION, picicons(45).Picture, picicons(45).Picture

    '------------------ WINDOW -------------
    'Cascade
    Load picicons(30)
    picicons(30).Picture = imlToolbarIcons.ListImages("Cascade").Picture
    picicons(30).Picture = picicons(30).Image

    SetMenuItemBitmaps hmnu(5), 0, MF_BYPOSITION, picicons(30).Picture, picicons(30).Picture

    'Tile Horizontally
    Load picicons(31)
    picicons(31).Picture = imlToolbarIcons.ListImages("TileHorizontally").Picture
    picicons(31).Picture = picicons(31).Image

    SetMenuItemBitmaps hmnu(5), 1, MF_BYPOSITION, picicons(31).Picture, picicons(31).Picture

    'Tile Vertically
    Load picicons(32)
    picicons(32).Picture = imlToolbarIcons.ListImages("TileVertically").Picture
    picicons(32).Picture = picicons(32).Image

    SetMenuItemBitmaps hmnu(5), 2, MF_BYPOSITION, picicons(32).Picture, picicons(32).Picture
    '--------------------------------

    '------------------ HELP -----------
    'Tips
    Load picicons(28)
    picicons(28).Picture = imlToolbarIcons.ListImages("Tip").Picture
    picicons(28).Picture = picicons(28).Image

    SetMenuItemBitmaps hmnu(6), 0, MF_BYPOSITION, picicons(28).Picture, picicons(28).Picture

    'Check for updates
    Load picicons(37)
    picicons(37).Picture = imlToolbarIcons.ListImages("CheckForUpdates").Picture
    picicons(37).Picture = picicons(37).Image

    SetMenuItemBitmaps hmnu(6), 2, MF_BYPOSITION, picicons(37).Picture, picicons(37).Picture

    '---------------------------------

'
'    On Error GoTo 0
'    Exit Sub
'
'SetIconsToMenus_Error:
'    HandleError Err, "frmGeneral.SetIconsToMenus"
End Sub

Private Sub LoadUpdateForm(Optional setfocus As Boolean = False)
'bring up browser
    On Error GoTo LoadUpdateForm_error
    
    If HaveComponent("msinet.ocx") Then
        Load frmCheckUpdate
        If setfocus Then frmCheckUpdate.setfocus

        If quickupdate Then Unload frmCheckUpdate
    Else
        If MessageBox("msinet.ocx is needed to automatically download DCME updates, do you want to visit our forums to download it manually?", vbYesNo + vbQuestion, "MSinet.ocx not found") = vbYes Then
            ShellExecute hWnd, "open", "http://forums.sscentral.com/index.php?showtopic=12845", _
                         vbNullString, vbNullString, SW_SHOWNORMAL
        End If
    End If
    
    On Error GoTo 0
    Exit Sub
    
LoadUpdateForm_error:
    'probably msinet is outdated, or something's wrong with it... disable auto update for now
    MessageBox "An error occured while loading the auto-update dialog. Auto-updates will be disabled for now. You can re-enable them from the Preferences... dialog." & vbCrLf & "Error " & Err.Number & " (" & Err.description & ")", vbExclamation + vbOKOnly
    
    Call SetSetting("AutoUpdate", 0)
    Call SaveSettings
    Exit Sub
End Sub

Function GetCTRL() As Boolean
    GetCTRL = Maps(activemap).usingctrl
End Function
Function GetShift() As Boolean
    GetShift = Maps(activemap).usingshift
End Function







Sub HideFloatTileset()
    Unload FloatTileset
    AutoHideTileset = False
End Sub

Private Sub ShowFloatTileset()
    FloatTileset.show
    AutoHideTileset = False
End Sub

Private Sub PopupTileset()
    dontUpdatePreview = True

    FloatTileset.Left = picRightBar.Left - FloatTileset.width + frmGeneral.Left + (3 * Screen.TwipsPerPixelX)
    FloatTileset.Top = frmGeneral.Top + toolbartop.height + picRightBar.Top + (picsmalltilepreview.Top * Screen.TwipsPerPixelY) + (TILEW * Screen.TwipsPerPixelY)
    FloatTileset.show
    AutoHideTileset = True
End Sub

Sub HideFloatRadar()
    Unload FloatRadar
    AutoHideRadar = False
End Sub

Private Sub ShowFloatRadar()
    FloatRadar.show
    AutoHideRadar = False
End Sub

Private Sub PopupRadar()
    dontUpdatePreview = True

    FloatRadar.width = picradar.width * Screen.TwipsPerPixelX
    FloatRadar.height = picradar.height * Screen.TwipsPerPixelY
    FloatRadar.Left = picRightBar.Left - FloatRadar.width + frmGeneral.Left + (3 * Screen.TwipsPerPixelX)
    FloatRadar.Top = frmGeneral.Top + toolbartop.height + picRightBar.Top + (picradar.Top * Screen.TwipsPerPixelY)
    FloatRadar.picradar.width = FloatRadar.ScaleWidth
    FloatRadar.picradar.height = FloatRadar.ScaleHeight

    FloatRadar.show
    AutoHideRadar = True
End Sub

Sub ChangeActiveRegionPythonCode(str As String)
    Call Maps(activemap).Regions.SetRegionPython(llRegionList.ListIndex, str)
End Sub

Function GetActiveRegionPythonCode() As String
    GetActiveRegionPythonCode = Maps(activemap).Regions.getRegionPython(llRegionList.ListIndex)
End Function

Function ActiveTilesetIs8bit() As Boolean
    ActiveTilesetIs8bit = Maps(activemap).TilesetIs8bit
End Function

Sub SetMousePointer(cursor As Integer)
    Screen.MousePointer = cursor
End Sub

Sub ResetMousePointer()
    Screen.MousePointer = 0 'vbDefault
End Sub

Function isTestmapActive(dontcheckonmap As frmMain) As Boolean
    Dim i As Integer
    For i = 0 To UBound(Maps)
        If loadedmaps(i) Then
            If Maps(i).TestMap.isRunning Then
                If Not Maps(i) Is dontcheckonmap Then
                    'Only return TRUE if testmap is active on a different map than the current one
                    isTestmapActive = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    isTestmapActive = False
End Function

Sub stopAllTestMap(dontcheckonmap As frmMain)
    Dim i As Integer
    For i = 0 To UBound(Maps)
        If loadedmaps(i) Then
            If Maps(i).TestMap.isRunning Then
                If Not Maps(i) Is dontcheckonmap Then
                    Maps(i).TestMap.StopRun
                End If
            End If
        End If
    Next
End Sub


'These functions should be used to get the directory where to put each item
'This could depend on user preferences
'Function folderLVZ(mapName As String, lvzname As String) As String
'10        folderLVZ = GetPathTo(mapName) & lvzname & "files"
'End Function
'
'Function folderCFG(mapName As String) As String
'
'End Function
'
'Function folderLVL(mapName As String) As String
'
'End Function

'''''

Sub UpdateAutoSaveList()
'Updates the File -> Autosaves menu

    Dim paths As String
    Dim splitpaths() As String
    Dim nextpath As String
    Dim i As Integer
    Dim j As Integer
    
    If Not loadedmaps(activemap) Then
        paths = ""
    Else
        paths = Dir$(App.path & "\DCME autosaves\" & Maps(activemap).eLVL.GetHashCode & "*.bak")
    End If

    For i = 1 To mnulstAutosaves.UBound
        Unload mnulstAutosaves(i)
    Next
    
    If paths = "" Then
        'no autosaves
        mnulstAutosaves(0).Caption = "No autosave found for this map"
        mnulstAutosaves(0).Enabled = False
        Exit Sub
    End If
    
    Do
        nextpath = Dir$()
        If nextpath <> "" Then
            paths = paths & "<>" & nextpath
        End If
    Loop While nextpath <> ""
    
    splitpaths = Split(paths, "<>")
    
    For i = 0 To UBound(splitpaths)
        
        If FileExists(App.path & "\DCME autosaves\" & splitpaths(i)) Then
            Dim mapInfo() As String
            mapInfo = Split(splitpaths(i), "_")
            '0: hashcode
            '1: date/time
            '2..: mapname
            If UBound(mapInfo) >= 2 Then
                If i <> 0 Then Load mnulstAutosaves(i)
                mnulstAutosaves(i).Enabled = True
            
                Dim mapTitle As String
                mapTitle = ""
                For j = 2 To UBound(mapInfo)
                    mapTitle = mapTitle & mapInfo(j) & IIf(j < UBound(mapInfo), "_", "")
                Next
                
                Dim mapDate() As String
                Dim mapDateString As String
                mapDate = Split(mapInfo(1), "-")
                If UBound(mapDate) = 5 Then
                    mapDateString = mapDate(0) & "/" & mapDate(1) & "/" & mapDate(2) & " " & mapDate(3) & ":" & mapDate(4) & ":" & mapDate(5)
                ElseIf UBound(mapDate) = 4 Then
                    mapDateString = mapDate(0) & "/" & mapDate(1) & "/" & mapDate(2) & " " & mapDate(3) & ":" & mapDate(4)
                Else
                    mapDateString = ""
                End If
                
                mnulstAutosaves(i).Caption = mapDateString & " " & Mid$(mapTitle, 1, Len(mapTitle) - 4)
                mnulstAutosaves(i).Tag = App.path & "\DCME autosaves\" & splitpaths(i)
            End If
        End If
    Next

End Sub

Public Property Let IsBusy(Optional ByVal RefMethod As String = vbNullString, ByVal newstatus As Boolean)
'Changes the IsBusy flag
'When set to true, disables frmGeneral completly
'When set to false, discards pending actions (DoEvents) and enables frmGeneral
'If a method is specified when setting it to true, the form will remain locked until the same method
'sets it to false.
'method name should be something like Line.MouseUp

    Static OwnerMethod As String
    
    If OwnerMethod = vbNullString Then
        OwnerMethod = RefMethod
        c_isbusy = newstatus
    
    ElseIf RefMethod <> vbNullString Then
        If OwnerMethod = RefMethod Then
            OwnerMethod = vbNullString
            c_isbusy = newstatus
        Else
            Exit Property
        End If
    Else
        Exit Property
    End If
    
    If c_isbusy = True Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
    Else
'180           DoEvents
        Me.Enabled = True
        Screen.MousePointer = MousePointerConstants.vbDefault
        

        Call HideProgress

    End If
End Property

Public Property Get IsBusy(Optional ByVal RefMethod As String = vbNullString) As Boolean
    IsBusy = c_isbusy
End Property

Sub ShowProgress(Operation As String, Optional Max As Long = 100)
    If Not c_showprogress Then
        'Load dlgProgress
    End If
    
    Call UpdateProgress(Operation, 0)
        
'50        Call dlgProgress.InitProgressBar(Operation, Max)
'60        dlgProgress.show vbModeless, Me
    
    c_showprogress = True
End Sub

Sub UpdateProgress(Operation As String, Optional value As Long = -1)
    Static ownerOperation As String
    Static lastupdate As Long
    
    If c_showprogress Then
        If Operation = ownerOperation Then
            
            progress = value
'40                Call dlgProgress.SetValue(progress)
            
            If GetTickCount - lastupdate > 200 Then
'50                    DoEvents
                lastupdate = GetTickCount
            End If
        End If
    Else
        ownerOperation = Operation
        If value <> -1 Then progress = value
    End If
    
End Sub

Public Sub UpdateProgressLabel(newLabel As String)
    If c_showprogress Then
'20            Call dlgProgress.SetLabel(newLabel)
'20            DoEvents
    End If
End Sub

Public Sub UpdateProgressOperation(Operation As String)
    If c_showprogress Then
        c_showprogress = False
        Call UpdateProgress(Operation, -1)
        c_showprogress = True
        
'50            Call dlgProgress.SetOperation(Operation)
    End If
End Sub

Sub HideProgress()
    If c_showprogress Then
        c_showprogress = False
'30            Unload dlgProgress
    End If
    
End Sub

Public Property Get progress() As Long
    progress = c_progress
End Property
Public Property Let progress(value As Long)
    c_progress = value
End Property


Function QuickSaveAll() As Boolean
'Saves every opened map that is mapchanged and appends "_recovery" to their name
'Used when a critical error occurs
    
    Dim i As Integer
    For i = 0 To 9
        If loadedmaps(i) Then
            If Maps(i).mapchanged Then
                Dim path As String
                If Maps(i).activeFile <> "" Then
                    path = GetPathTo(Maps(i).activeFile) & GetFileTitle(Maps(i).activeFile) & "_recovery.lvl"
                Else
                    path = App.path & "\" & Maps(i).Caption & "_recovery.lvl"
                End If
                
                Call Maps(i).SaveMap(path, (SFdefault Or SFsilent))
            End If
        End If
    Next
    Unload Me
End Function


Sub RestoreWindow()
    On Error GoTo RestoreWindow_Error
    
   Dim currWinP As WINDOWPLACEMENT
   
  'if a window handle passed
    If Me.hWnd Then
    
    'prepare the WINDOWPLACEMENT type
    'to receive the window coordinates
    'of the specified handle
        currWinP.Length = Len(currWinP)
        
        'get the info...
        If GetWindowPlacement(Me.hWnd, currWinP) > 0 Then
        'based on the returned info,
        'determine the window state
            If currWinP.showCmd = SW_SHOWMINIMIZED Then
                'it is minimized, so restore it
                With currWinP
                    .Length = Len(currWinP)
                    .flags = 0&
                    If lastwindowstate = vbMaximized Then
                        .showCmd = SW_SHOWMAXIMIZED
                    Else
                        .showCmd = SW_SHOWDEFAULT
                    End If
                End With
                
                Call SetWindowPlacement(Me.hWnd, currWinP)
            Else
                'it is on-screen, so make it visible
                Call SetForegroundWindow(Me.hWnd)
                Call BringWindowToTop(Me.hWnd)
            End If
        
        End If
    End If
   
   Exit Sub
   
RestoreWindow_Error:
    HandleError Err, "RestoreWindow"
End Sub


Private Sub LoadRevert()
    On Error GoTo LoadRevert_Error
    
    If Not loadedmaps(activemap) Then Exit Sub
    
'          Dim revpath As String
    Dim curpath As String
    
'30        revpath = Maps(activemap).RevertPath
    curpath = Maps(activemap).activeFile
    
'50        If Not FileExists(revpath) Then
'60             messagebox "No revert available for this map. " & revpath & " not found."
'70             Exit Sub
'80        End If
    If curpath = "" Then
      MessageBox "Map was not saved yet.", vbExclamation
      Exit Sub
    ElseIf Not FileExists(curpath) Then
      MessageBox "Map file '" & curpath & "' not found!", vbExclamation
      Exit Sub
    End If
    
    IsBusy("frmGeneral.LoadRevert") = True
    
    'Trigger autosave of current map
    Call Maps(activemap).DoAutoSave(True)
  
    Call DestroyMap(activemap)
'140       DoEvents
    Call OpenMap(curpath)
    
    Call UpdateMenuMaps
    
'110       If curpath <> "" Then
'120         If DeleteFile(curpath) Then
'130               FileCopy revpath, curpath
'140         Else
'                'Could not delete the current file for some reason
'                'We'll open the revert from its current location
'150             curpath = revpath
'160         End If
'170       Else
'             'No path specified, we'll just open the revert from its
'             'current location
'180          curpath = revpath
'190       End If
'
'200       If FileExists(curpath) Then
'
'              'We don't want it to ask to save
'210           Maps(activemap).mapchanged = False
'
'220           Call DestroyMap(activemap)
'230           DoEvents
'240           Call OpenMap(curpath)
'
'250           Maps(activemap).mapchanged = False
'
'260           Call UpdateToolBarButtons
'270           Call UpdateMenuMaps
'
'280       End If
    
    IsBusy("frmGeneral.LoadRevert") = False
    
    On Error GoTo 0
    Exit Sub
LoadRevert_Error:
      IsBusy("frmGeneral.LoadRevert") = False
    HandleError Err, "frmGeneral.LoadRevert"
End Sub


Private Sub txtLVZDisplayTime_Change()
    If loadedmaps(activemap) And txtLVZDisplayTime.Text <> "" Then
        Call removeDisallowedCharacters(txtLVZDisplayTime, 0, 4095, False)
        Call Maps(activemap).lvz.ChangeSelectionDisplayTime(CLng(txtLVZDisplayTime.Text))
    End If
End Sub

Private Sub txtLvzObjectID_Change()
    If loadedmaps(activemap) And txtLvzObjectID <> "" Then
        Call removeDisallowedCharacters(txtLvzObjectID, 0, 32767, False)
        
        Call Maps(activemap).lvz.ChangeSelectionObjectID(CInt(txtLvzObjectID.Text))
    End If
End Sub

Private Sub txtLvzSnap_Change()
    If loadedmaps(activemap) Then

        Call removeDisallowedCharacters(txtLvzSnap, 1, 512, False)
        Maps(activemap).lvz.snap = CInt(val(txtLvzSnap.Text))

    End If
End Sub

Private Sub txtLvzX_Change()
    If loadedmaps(activemap) And txtLvzX <> "" And txtLvzX.Enabled Then
        If Not Maps(activemap).lvz.dontCallSelectionChange Then
            Call removeDisallowedCharacters(txtLvzX, -32768, 32767, False)
        
            Call Maps(activemap).lvz.ChangeSelectionX(CInt(txtLvzX.Text))
        End If
    End If
End Sub

Private Sub txtLvzY_Change()
    If loadedmaps(activemap) And txtLvzY <> "" And txtLvzY.Enabled Then
        If Not Maps(activemap).lvz.dontCallSelectionChange Then
            Call removeDisallowedCharacters(txtLvzY, -32768, 32767, False)
        
            Call Maps(activemap).lvz.ChangeSelectionY(CInt(txtLvzY.Text))
        End If
    End If
End Sub




Friend Function GetLayer(layer As Integer) As clsDisplayLayer
    Set GetLayer = MapLayers(layer)
End Function

Public Property Get LeftTilesetColor() As Long

    LeftTilesetColor = m_lLeftTilesetColor

End Property

Public Property Let LeftTilesetColor(ByVal lLeftTilesetColor As Long)

    m_lLeftTilesetColor = lLeftTilesetColor

    cTileset.LeftColor = m_lLeftTilesetColor
End Property

Public Property Get RightTilesetColor() As Long

    RightTilesetColor = m_lRightTilesetColor

    cTileset.RightColor = m_lRightTilesetColor
End Property

Public Property Let RightTilesetColor(ByVal lRightTilesetColor As Long)

    m_lRightTilesetColor = lRightTilesetColor

End Property

Public Property Get TilesetBackgroundColor() As Long

    TilesetBackgroundColor = m_lTilesetBackgroundColor

End Property

Public Property Let TilesetBackgroundColor(ByVal lTilesetBackgroundColor As Long)

    m_lTilesetBackgroundColor = lTilesetBackgroundColor

End Property
