VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7020
   ClientLeft      =   8790
   ClientTop       =   3630
   ClientWidth     =   8865
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCleanup 
      Caption         =   "Cleanup"
      Height          =   375
      Left            =   1800
      TabIndex        =   116
      ToolTipText     =   "Cleans up the unused settings in the settings.dat file"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5940
      Index           =   4
      Left            =   8640
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   563
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   0
      Width           =   8445
      Begin VB.ComboBox cmbImageEditor 
         Height          =   315
         ItemData        =   "frmOptions.frx":000C
         Left            =   4080
         List            =   "frmOptions.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   1560
         Width           =   2055
      End
      Begin VB.PictureBox picIconImageEditor 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   105
         Top             =   1440
         Width           =   480
      End
      Begin VB.Frame Frame8 
         Caption         =   "Image Editing"
         Height          =   855
         Left            =   120
         TabIndex        =   101
         Top             =   1200
         Width           =   7815
         Begin VB.CommandButton cmdBrowseImageEditor 
            Caption         =   "Browse..."
            Height          =   495
            Left            =   6120
            TabIndex        =   103
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Use the following program to edit images :"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "LVL files"
         Height          =   825
         Left            =   120
         TabIndex        =   99
         Top             =   120
         Width           =   7815
         Begin VB.CommandButton cmdAssociateLVL 
            Caption         =   "Associate LVL Files"
            Height          =   495
            Left            =   6120
            TabIndex        =   100
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblLVLAssociation 
            Caption         =   "Label6"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   4455
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4620
      Index           =   1
      Left            =   480
      ScaleHeight     =   4620
      ScaleWidth      =   8445
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7500
      Width           =   8445
      Begin VB.Frame frameGrid 
         Caption         =   "Colors"
         Height          =   4425
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2415
         Begin VB.Frame Frame1 
            Caption         =   "Horizontal Lines"
            Height          =   1815
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2055
            Begin VB.Label lblGridColorsA 
               Caption         =   "Center Line:"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Section Lines:"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   30
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Block Lines:"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   32
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Grid Lines:"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   34
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   1680
               MousePointer    =   2  'Cross
               TabIndex        =   29
               Top             =   360
               Width           =   255
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   1
               Left            =   1680
               MousePointer    =   2  'Cross
               TabIndex        =   31
               Top             =   720
               Width           =   255
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   2
               Left            =   1680
               MousePointer    =   2  'Cross
               TabIndex        =   33
               Top             =   1080
               Width           =   255
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   3
               Left            =   1680
               MousePointer    =   2  'Cross
               TabIndex        =   35
               Top             =   1440
               Width           =   255
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Vertical Lines"
            Height          =   2055
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   2055
            Begin VB.CheckBox chkColorY 
               Caption         =   "Same As Horizontal"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Grid Lines:"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   44
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Block Lines:"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   42
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Section Lines:"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   40
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblGridColorsA 
               Caption         =   "Center Line:"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   38
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   4
               Left            =   1680
               TabIndex        =   39
               Top             =   600
               Width           =   255
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   5
               Left            =   1680
               TabIndex        =   41
               Top             =   960
               Width           =   255
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   6
               Left            =   1680
               TabIndex        =   43
               Top             =   1320
               Width           =   255
            End
            Begin VB.Label lbl_color 
               BackColor       =   &H80000007&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   7
               Left            =   1680
               TabIndex        =   45
               Top             =   1680
               Width           =   255
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Spacing"
         Height          =   4425
         Left            =   2520
         TabIndex        =   46
         Top             =   0
         Width           =   5775
         Begin VB.Frame frmSpaceY 
            Caption         =   "Vertical Lines"
            Height          =   2055
            Left            =   120
            TabIndex        =   58
            Top             =   2160
            Width           =   5535
            Begin VB.CheckBox chkSpaceY 
               Caption         =   "Same as Horizontal Lines"
               Height          =   255
               Left            =   240
               TabIndex        =   59
               Top             =   240
               Width           =   2175
            End
            Begin VB.CommandButton cmd_centeroffset 
               Caption         =   "Center With Map"
               Height          =   495
               Index           =   1
               Left            =   3840
               TabIndex        =   65
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtBlocks 
               Height          =   285
               Index           =   1
               Left            =   1680
               MaxLength       =   3
               TabIndex        =   60
               Text            =   "32"
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txtSections 
               Height          =   285
               Index           =   1
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   68
               Text            =   "4"
               Top             =   1560
               Width           =   375
            End
            Begin MSComctlLib.Slider GridOffset 
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   63
               Top             =   1080
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   1
               Max             =   30
            End
            Begin VB.Label lbl_offset 
               Caption         =   "0"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   66
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label lblSpace 
               Caption         =   "Offset:"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   64
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lblSpace 
               Caption         =   "Section Lines Every:"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   67
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label lblSpace 
               Caption         =   "Block Lines Every:"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   62
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label lblSpace 
               Caption         =   "Tiles"
               Height          =   255
               Index           =   6
               Left            =   2160
               TabIndex        =   61
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblSpace 
               Caption         =   "Blocks"
               Height          =   255
               Index           =   5
               Left            =   2160
               TabIndex        =   69
               Top             =   1560
               Width           =   855
            End
         End
         Begin VB.Frame frmSpaceX 
            Caption         =   "Horizontal Lines"
            Height          =   1815
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   5535
            Begin VB.TextBox txtSections 
               Height          =   285
               Index           =   0
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   56
               Text            =   "4"
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox txtBlocks 
               Height          =   285
               Index           =   0
               Left            =   1680
               MaxLength       =   3
               TabIndex        =   49
               Text            =   "32"
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton cmd_centeroffset 
               Caption         =   "Center With Map"
               Height          =   495
               Index           =   0
               Left            =   3840
               TabIndex        =   53
               Top             =   720
               Width           =   1455
            End
            Begin MSComctlLib.Slider GridOffset 
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   51
               Top             =   840
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   1
               Max             =   30
            End
            Begin VB.Label lblSpace 
               Caption         =   "Blocks"
               Height          =   255
               Index           =   4
               Left            =   2160
               TabIndex        =   57
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label lblSpace 
               Caption         =   "Tiles"
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   50
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblSpace 
               Caption         =   "Block Lines Every:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   48
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label lblSpace 
               Caption         =   "Section Lines Every:"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   55
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label lblSpace 
               Caption         =   "Offset:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   52
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label lbl_offset 
               Caption         =   "0"
               Height          =   255
               Index           =   0
               Left            =   3240
               TabIndex        =   54
               Top             =   840
               Width           =   855
            End
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5940
      Index           =   3
      Left            =   8520
      ScaleHeight     =   5940
      ScaleWidth      =   8445
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   720
      Width           =   8445
      Begin VB.Frame Frame3 
         Caption         =   "Automatic Updates Settings"
         Height          =   1425
         Left            =   120
         TabIndex        =   76
         Top             =   120
         Width           =   5895
         Begin VB.CheckBox chkUpdateURLCustom 
            Caption         =   "Get Update Information From Custom URL"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   600
            Width           =   5415
         End
         Begin VB.TextBox txtUpdateURL 
            Height          =   285
            Left            =   240
            TabIndex        =   84
            Text            =   "url"
            Top             =   960
            Width           =   5415
         End
         Begin VB.CheckBox chkAutoUpdate 
            Caption         =   "Check for updates"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbUpdateDelay 
            Height          =   315
            ItemData        =   "frmOptions.frx":0047
            Left            =   2040
            List            =   "frmOptions.frx":0057
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.PictureBox piccurrenttileset 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   9000
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   0
      Top             =   -1200
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox picWalltiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   9000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Reset To Default"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox picdefaulttileset 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   8880
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5940
      Index           =   2
      Left            =   240
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   563
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   480
      Width           =   8445
      Begin VB.Frame frmDispTileset 
         Caption         =   "Tileset"
         Height          =   1335
         Left            =   120
         TabIndex        =   107
         Top             =   3840
         Width           =   8055
         Begin VB.Frame Frame9 
            Caption         =   "Preview"
            Height          =   1095
            Left            =   6360
            TabIndex        =   112
            Top             =   120
            Width           =   1575
            Begin VB.PictureBox picPreviewTileset 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   780
               Left            =   120
               ScaleHeight     =   52
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   92
               TabIndex        =   113
               Top             =   240
               Width           =   1380
               Begin VB.Shape shpPreviewTileset 
                  Height          =   510
                  Left            =   0
                  Top             =   0
                  Width           =   510
               End
            End
         End
         Begin VB.Label lblLeftColor 
            Caption         =   "Left Selection Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblRightColor 
            Caption         =   "Right Selection Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblTilesetBackground_color 
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3120
            MousePointer    =   2  'Cross
            TabIndex        =   111
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblRightColor_color 
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3120
            MousePointer    =   2  'Cross
            TabIndex        =   110
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblTilesetPreviewGBColor 
            Caption         =   "Background Color Under Tile Preview:"
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label lblLeftColor_color 
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3120
            MousePointer    =   2  'Cross
            TabIndex        =   108
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame frmRegions 
         Caption         =   "ASSS Regions"
         Height          =   855
         Left            =   120
         TabIndex        =   93
         Top             =   2880
         Width           =   8055
         Begin DCME.cUpDown udRegnOpacity2 
            Height          =   375
            Left            =   6360
            TabIndex        =   95
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Max             =   255
         End
         Begin DCME.cUpDown udRegnOpacity 
            Height          =   375
            Left            =   2520
            TabIndex        =   94
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Max             =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Opacity (Other Tool Selected)"
            Height          =   255
            Left            =   4080
            TabIndex        =   97
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Opacity (Region Tool Selected)"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "LVZ"
         Height          =   735
         Left            =   120
         TabIndex        =   91
         Top             =   2040
         Width           =   8055
         Begin VB.CheckBox chkLVZImagesAnimatedTiles 
            Caption         =   "Use animation for LVZ Image Library"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Radar"
         Height          =   735
         Left            =   120
         TabIndex        =   89
         Top             =   1200
         Width           =   8055
         Begin VB.CheckBox chkAutoFullMapPreview 
            Caption         =   "Auto switch to full map preview (Can reduce performance)"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.Frame frmCursor 
         Caption         =   "Cursor"
         Height          =   945
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   8055
         Begin VB.CheckBox chkShowTileCoord 
            Caption         =   "Show Tile Coordinates"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkTilePreview 
            Caption         =   "Show Tile Preview Under Cursor"
            Height          =   255
            Left            =   4080
            TabIndex        =   72
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lbl_color 
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   8
            Left            =   1440
            MousePointer    =   2  'Cross
            TabIndex        =   74
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblGridColorsA 
            Caption         =   "Cursor Color:"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   73
            Top             =   600
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5940
      Index           =   0
      Left            =   7200
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   563
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   8445
      Begin VB.Frame frmAutosave 
         Caption         =   "Autosave Settings"
         Height          =   1095
         Left            =   120
         TabIndex        =   79
         Top             =   1200
         Width           =   8175
         Begin VB.CheckBox chkAutosave 
            Caption         =   "Autosave every"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   240
            Width           =   1575
         End
         Begin DCME.cUpDown UpDownAutosave 
            Height          =   300
            Left            =   1920
            TabIndex        =   86
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Min             =   1
            Max             =   120
            Value           =   2
         End
         Begin DCME.cUpDown updownMaxAutosaves 
            Height          =   300
            Left            =   1200
            TabIndex        =   87
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Min             =   1
            Max             =   99
            Value           =   2
         End
         Begin VB.Label Label1 
            Caption         =   "minutes"
            Height          =   255
            Left            =   2760
            TabIndex        =   83
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "autosaves of each map"
            Height          =   255
            Left            =   2040
            TabIndex        =   82
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblMaxAutosave 
            Caption         =   "Keep up to"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame frmWalltiles 
         Caption         =   "Walltiles"
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   8175
         Begin VB.PictureBox picWallTilesPrev 
            AutoRedraw      =   -1  'True
            Height          =   1470
            Left            =   5160
            ScaleHeight     =   94
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   188
            TabIndex        =   18
            Top             =   120
            Width           =   2880
         End
         Begin VB.OptionButton optDefaultWall 
            Caption         =   "Specify Default Walltiles Set"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton optDefaultWall 
            Caption         =   "Use Empty Walltiles Set"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmdBrowseWalltiles 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   4200
            TabIndex        =   16
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtPathWall 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   960
            Width           =   4815
         End
      End
      Begin VB.Frame frmGeneralSettings 
         Caption         =   "General Settings"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8175
         Begin VB.CheckBox chkNewMap 
            Caption         =   "Create new map at startup"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkTips 
            Caption         =   "Show tips at startup"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame frmTileset 
         Caption         =   "Tileset"
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   8175
         Begin VB.CommandButton cmdImport 
            Caption         =   "Import Now"
            Height          =   255
            Left            =   3840
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox picDefaultPreview 
            AutoRedraw      =   -1  'True
            Height          =   1470
            Left            =   5160
            ScaleHeight     =   94
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   178
            TabIndex        =   12
            Top             =   120
            Width           =   2730
         End
         Begin VB.OptionButton optDefaultTileset 
            Caption         =   "Specify Default Tileset"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton optDefaultTileset 
            Caption         =   "Use Standard Default Tileset"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdBrowseTileset 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   3840
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   960
            Width           =   4815
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   6405
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11298
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grid"
            Key             =   "Grid"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Display"
            Key             =   "Display"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Updates"
            Key             =   "Updates"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "File Associations"
            Key             =   "File Associations"
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   13080
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tilesetpath As String
Dim wallsetpath As String
Dim imageeditorpath As String


Dim walltiles As New walltiles

Private Sub chkAutoUpdate_Click()
        cmbUpdateDelay.Enabled = (chkAutoUpdate.value = vbChecked)
End Sub

Private Sub chkColorY_Click()
    Call CheckColorY
End Sub





Private Sub chkSpaceY_Click()
    Call CheckSpaceY
End Sub

Private Sub CheckSpaceY()
    Dim i As Integer

    For i = 5 To 9
        lblSpace(i).Enabled = chkSpaceY.value
    Next
    
    txtBlocks(1).Enabled = chkSpaceY.value
    txtSections(1).Enabled = chkSpaceY.value
    GridOffset(1).Enabled = chkSpaceY.value
    lbl_offset(1).Enabled = chkSpaceY.value
    
    If chkSpaceY.value = vbChecked Then

        txtBlocks(1).Text = txtBlocks(0).Text
        txtSections(1).Text = txtSections(0).Text
        GridOffset(1).value = GridOffset(0).value

    Else

        txtBlocks(1).Text = GetSetting("GridBlocksY", "51")
        txtSections(1).Text = GetSetting("GridSectionsY", "4")
        GridOffset(1).value = GetSetting("GridOffsetY", "0")

    End If
End Sub




Private Sub chkUpdateURLCustom_Click()
    Static customURL As String
    If chkUpdateURLCustom.value = vbChecked Then
        txtUpdateURL.Enabled = True
        If customURL <> "" Then
            txtUpdateURL.Text = customURL
        End If
    Else
        txtUpdateURL.Enabled = False
        customURL = txtUpdateURL.Text
        txtUpdateURL.Text = DEFAULT_UPDATE_URL
    End If
End Sub

Private Sub cmbImageEditor_Click()
    Select Case cmbImageEditor.ListIndex
    Case 0
        'Current one that is already selected
        
    Case 1
        'System's default
        Call SetImageEditor(GetSystemImageEditor)
    Case 2
        'MS paint
        Call SetImageEditor(SysDir & "\mspaint.exe", "MS Paint")
    End Select
End Sub

Private Sub cmdAssociateLVL_Click()
    If Not bDEBUG Then Call AssignExt
End Sub

Private Sub cmdBrowseImageEditor_Click()
    On Error GoTo errh

    cd.DialogTitle = "Open images with..."
    cd.flags = cdlOFNHideReadOnly
    cd.InitDir = "C:\Program Files"
    cd.Filter = "*.exe|*.exe"

    cd.ShowOpen

    If cd.filename <> "" And FileExists(cd.filename) Then
        Call SetImageEditor(cd.filename)
    End If
    
    cd.InitDir = App.path
    
    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
    End If
End Sub

Private Sub cmdBrowseWalltiles_Click()
    On Error GoTo errh

    cd.DialogTitle = "Load walltiles..."
    cd.flags = cdlOFNHideReadOnly
    cd.filename = "Default_walls.wtl"
    cd.Filter = "*.wtl|*.wtl"

    cd.ShowOpen

    If cd.filename <> "" Then
        Call walltiles.LoadWallTiles(cd.filename)
        wallsetpath = cd.filename

        Call SetDefaultWall(wallsetpath, True)
    End If

    Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
    End If
End Sub

Private Sub cmdApply_Click()
    SavePreferences
End Sub

Private Sub cmdBrowseTileset_Click()
    Call BrowseTileset

    Call SetDefaultTileset(txtPath.Text)

    Call UpdatePreview
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmd_centeroffset_Click(Index As Integer)
    GridOffset(Index).value = 512 Mod txtBlocks(Index).Text
End Sub

Private Sub cmdCleanup_Click()
    Call ClearSettings
    Call SaveSettings
End Sub

Private Sub cmdDefault_Click()
    If MessageBox("Reset all settings to default?", vbYesNo + vbExclamation, "Reset settings") = vbYes Then
        Call ClearSettings
        Call LoadSettings(True)
    End If
End Sub

Private Sub cmdImport_Click()
    If FileExists(txtPath.Text) Then
        Call frmGeneral.ApplyEditedTileset(txtPath.Text)
    Else
        MessageBox txtPath.Text & " not found.", vbExclamation + vbOKOnly, "File not found"
    End If
End Sub

Private Sub cmdOK_Click()
    If GetExtension(txtUpdateURL.Text) <> "txt" And GetExtension(txtUpdateURL.Text) <> "ini" Then
        Call txtUpdateURL_LostFocus
        Exit Sub
    End If
    
    SavePreferences
    Unload Me
End Sub





Private Sub Form_Activate()
    DoEvents
    Call UpdatePreview
End Sub

Private Sub Form_Load()
'center the form
    Me.Move (Screen.width - Me.width) / 2, (Screen.Height - Me.Height) / 2
    Call tbsOptions_Click
    Set Me.Icon = frmGeneral.Icon

    Call LoadSettings
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub





Private Sub lbl_color_Click(Index As Integer)
    Dim defaultColor As Long
    Select Case Index
        Case 0, 4
            defaultColor = RGB(0, 0, 200)
        Case 1, 5
            defaultColor = RGB(200, 0, 0)
        Case 2, 6
            defaultColor = RGB(150, 150, 150)
        Case 3, 7
            defaultColor = RGB(80, 80, 80)
        Case 8
            defaultColor = RGB(230, 230, 230)
        Case Else
            defaultColor = vbBlack
    End Select
    
    lbl_color(Index).BackColor = GetColor(Me, lbl_color(Index).BackColor, False, True, defaultColor)
    Call CheckColorY
End Sub





Private Sub lblLeftColor_color_Click()
    lblLeftColor_color.BackColor = GetColor(Me, lblLeftColor_color.BackColor, False, True, DEFAULT_LEFTCOLOR)
    
    Call DrawTilesetSettingsPreview
End Sub

Private Sub lblRightColor_color_Click()
    lblRightColor_color.BackColor = GetColor(Me, lblRightColor_color.BackColor, False, True, DEFAULT_RIGHTCOLOR)
    
    Call DrawTilesetSettingsPreview
End Sub


Private Sub lblTilesetBackground_color_Click()
    lblTilesetBackground_color.BackColor = GetColor(Me, lblTilesetBackground_color.BackColor, True, True, DEFAULT_TILESETBACKGROUND)
    
    Call DrawTilesetSettingsPreview
End Sub

Private Sub optDefaultTileset_Click(Index As Integer)
    If Index = 0 Then
        optDefaultTileset(0).value = True
        optDefaultTileset(1).value = False
        Call SetDefaultTileset("None")
    Else
        optDefaultTileset(1).value = True
        optDefaultTileset(0).value = False
        cmdBrowseTileset.Enabled = True
        Call SetDefaultTileset(tilesetpath)
    End If

    DoEvents
    Call UpdatePreview

End Sub

Private Sub optDefaultWall_Click(Index As Integer)
    If Index = 0 Then
        optDefaultWall(0).value = True
        optDefaultWall(1).value = False
        Call SetDefaultWall("", False)
    Else
        optDefaultWall(1).value = True
        optDefaultWall(0).value = False
        Call SetDefaultWall(wallsetpath, True)
    End If

    DoEvents
    Call DrawWallTiles
End Sub





Private Sub tbsOptions_Click()

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 15
            picOptions(i).Enabled = True
            picOptions(i).Top = 32
        Else
            picOptions(i).Left = -5000
            picOptions(i).Enabled = False
        End If
    Next

    If tbsOptions.SelectedItem.Index = 1 Then
        DoEvents
        Call UpdatePreview
    End If

End Sub


Private Sub txtBlocks_Change(Index As Integer)
    Call removeDisallowedCharacters(txtBlocks(Index), 1, 99999)

    If GridOffset(Index).value > val(txtBlocks(Index).Text) Then GridOffset(Index).value = val(txtBlocks(Index).Text) - 1
    If val(txtBlocks(Index).Text) > 1 Then
        GridOffset(Index).Max = val(txtBlocks(Index).Text) - 1
    Else
        GridOffset(Index).Max = 1
    End If
    GridOffset_Change (Index)

    If chkSpaceY.value = vbChecked Then txtBlocks(1).Text = txtBlocks(0).Text
End Sub


Private Sub txtSections_Change(Index As Integer)
    Call removeDisallowedCharacters(txtBlocks(Index), 1, 99999)

    If val(txtSections(Index).Text) <= 0 Then
        txtSections(Index).Text = 1
    End If

    If val(txtSections(Index).Text) >= val(txtBlocks(Index).Text) Then
        txtSections(Index).Text = val(txtBlocks(Index).Text)
    End If

    If chkSpaceY.value = vbChecked Then txtSections(1).Text = txtSections(0).Text
End Sub

Private Sub GridOffset_Change(Index As Integer)
    If GridOffset(Index).value = 1 Then
        lbl_offset(Index).Caption = GridOffset(Index).value & " Tile"
    Else
        lbl_offset(Index).Caption = GridOffset(Index).value & " Tiles"
    End If

    If chkSpaceY.value = vbChecked Then GridOffset(1).value = GridOffset(0).value

End Sub

Sub SavePreferences()
    Dim i As Integer

    Call SetSetting("GridBlocksX", txtBlocks(0).Text)
    Call SetSetting("GridSectionsX", txtSections(0).Text)
    Call SetSetting("GridOffsetX", CStr(GridOffset(0).value))

    i = IIf(chkSpaceY.value = vbChecked, 0, 1)
    Call SetSetting("GridBlocksY", txtBlocks(i).Text)
    Call SetSetting("GridSectionsY", txtSections(i).Text)
    Call SetSetting("GridOffsetY", CStr(GridOffset(i).value))

    For i = 0 To 7
        Call SetSetting("GridColor" & CStr(i), CStr(lbl_color(i).BackColor))
    Next

    If txtPath.Text = "None" Or txtPath.Text = "" Then
        Call SetSetting("DefaultTileset", vbNullString)
    Else
        Call SetSetting("DefaultTileset", txtPath.Text)
    End If

    If txtPathWall.Text = "None" Or txtPathWall.Text = "" Then
        Call SetSetting("DefaultWalltiles", vbNullString)
    Else
        Call SetSetting("DefaultWalltiles", txtPathWall.Text)
    End If

    Call SetSetting("ShowTips", chkTips.value)

    
    Call SetSetting("AutoUpdate", chkAutoUpdate.value)
    Call SetSetting("AutoUpdateDelay", cmbUpdateDelay.ListIndex)
    If GetSetting("UpdateURL", DEFAULT_UPDATE_URL) <> DEFAULT_UPDATE_URL Or _
        txtUpdateURL.Text <> DEFAULT_UPDATE_URL Then
        Call SetSetting("UpdateURL", txtUpdateURL.Text)
    End If
    
    Call SetSetting("AutoNewMap", chkNewMap.value)

    Call SetSetting("ShowCursorCoords", chkShowTileCoord.value)
    Call SetSetting("CursorColor", CStr(lbl_color(8).BackColor))
    Call SetSetting("ShowPreview", chkTilePreview.value)
    

    Call SetSetting("RegnOpacity1", udRegnOpacity.value)
    Call SetSetting("RegnOpacity2", udRegnOpacity2.value)
'    Call SetSetting("RegnDrawTopOnly", chkOnlyTopRegion.value)
        
    Call SetSetting("MaxAutosaves", CStr(updownMaxAutosaves.value))
    Call SetSetting("AutoSaveDelay", CStr(UpDownAutosave.value))
    Call SetSetting("AutoSaveEnable", chkAutosave.value)

    Call SetSetting("AutoFullMapPreview", chkAutoFullMapPreview.value)
    Call SetSetting("AnimatedLVZImageTiles", chkLVZImagesAnimatedTiles.value)
    
    
    Call SetSetting("LeftColor", lblLeftColor_color.BackColor)
    Call SetSetting("RightColor", lblRightColor_color.BackColor)
    Call SetSetting("TilesetBackground", lblTilesetBackground_color.BackColor)
    
    
    
    Call SetSetting("ImageEditor", imageeditorpath)
    
    
    Call frmGeneral.ExecuteUpdate(True)

    Call settings.SaveSettings
End Sub

Private Sub CheckColorY()
    Dim i As Integer

    If chkColorY.value = vbUnchecked Then
        For i = 4 To 7
            lbl_color(i).Enabled = True
            lblGridColorsA(i).Enabled = True
            lbl_color(i).MousePointer = 2
        Next
    Else
        For i = 4 To 7
            lbl_color(i).Enabled = False
            lbl_color(i).BackColor = lbl_color(i - 4).BackColor
            lblGridColorsA(i).Enabled = False
            lbl_color(i).MousePointer = 12
        Next
    End If
End Sub

Sub SetColor(Index As Integer, color As Long)
    lbl_color(Index).BackColor = color
    If Index <= 3 And chkColorY.value = vbChecked Then
        lbl_color(Index + 4).BackColor = color
    End If
End Sub

Private Sub BrowseTileset()
'Specify a tileset
    On Error GoTo errh

    'opens a common dialog
    cd.DialogTitle = "Select an external default tileset"
    cd.flags = cdlOFNHideReadOnly
    cd.Filter = "*.lvl, *.bmp|*.lvl; *.bmp"
    cd.ShowOpen

    If GetExtension(cd.filetitle) = "lvl" Then
        If Not frmGeneral.hasLVLaTileset(cd.filename) Then
            'load default tileset
            MessageBox "This map does not contain a tileset.", vbOKOnly, "Tileset not found"
            Exit Sub
        End If
    End If

    'Open the file
    Dim f As Integer

    f = FreeFile

    Dim bm As Integer
    Dim Size As Long

    Dim b() As Byte

    Dim tmpfileh As BITMAPFILEHEADER
    Dim tmpinfoh As BITMAPINFOHEADER

    Open cd.filename For Binary As #f
    ReDim b(1)
    Get #f, , b

    If Chr(b(0)) & Chr(b(1)) <> "BM" Then
        'if there is no BM it means that we have either a
        'corrupt lvl or bmp file, or we have a lvl file
        'that uses the default tileset
    Else
        'rewind to beginning
        Seek #f, 1
        Get #f, , tmpfileh
        Get #f, , tmpinfoh
        'check if its lower than 8 bit
        'and check for width=304 and height=160
        'if it's not, we have an invalid bmp or lvl
        'to import
        If tmpinfoh.biBitCount < 8 Or tmpinfoh.biHeight <> 160 Or tmpinfoh.biWidth <> 304 Then
            MessageBox "Tilesets are required to be 304x160 pixels, with 8-bit or 24-bit color depth!", vbExclamation
            Close #f
            Exit Sub
        End If
    End If
    Close #f

    picdefaulttileset.Picture = LoadPicture(cd.filename)
    picdefaulttileset.Refresh
    txtPath.Text = cd.filename
    cmdImport.Enabled = True
    tilesetpath = txtPath.Text

    Exit Sub
errh:
    'if something goes wrong, use the default tileset
    If Err = cdlCancel Then
        'Ignore error
    Else
        MessageBox Err & " " & Err.description, vbCritical
    End If

    On Error GoTo 0
    Exit Sub


End Sub

Private Sub SetDefaultTileset(path As String)

    txtPath.Text = path
    
    cmdBrowseTileset.Enabled = (path <> "None")
    
    Call UpdatePreview

End Sub

Private Sub SetDefaultWall(path As String, enablebrowse As Boolean)

    txtPathWall.Text = IIf(path = "", "None", path)
    
    cmdBrowseWalltiles.Enabled = enablebrowse

    Call DrawWallTiles
End Sub

Private Sub UpdatePreview()


    If optDefaultTileset(0).value = True Then
        BitBlt piccurrenttileset.hDc, 0, 0, picdefaulttileset.width, picdefaulttileset.Height, frmGeneral.picdefaulttileset.hDc, 0, 0, vbSrcCopy
        piccurrenttileset.Refresh
        cmdImport.Enabled = False
    Else
        BitBlt piccurrenttileset.hDc, 0, 0, picdefaulttileset.width, picdefaulttileset.Height, picdefaulttileset.hDc, 0, 0, vbSrcCopy
        piccurrenttileset.Refresh
        cmdImport.Enabled = True
    End If

    SetStretchBltMode picDefaultPreview.hDc, HALFTONE
    StretchBlt picDefaultPreview.hDc, 0, 0, picDefaultPreview.ScaleWidth, picDefaultPreview.ScaleHeight, piccurrenttileset.hDc, 0, 0, picdefaulttileset.width, picdefaulttileset.Height, vbSrcCopy

    picDefaultPreview.Refresh
    Call DrawWallTiles
    
    Call DrawTilesetSettingsPreview
End Sub

Private Sub LoadSettings(Optional reset As Boolean = False)
'''load the settings into the text boxes
    Dim ret As String
'    Dim ret2 As String
    Dim i As Integer

    'Block lines frequency
    txtBlocks(0).Text = LoadSetting("GridBlocksX", DEFAULT_GRID_BLOCKS, reset)
    txtBlocks(1).Text = LoadSetting("GridBlocksY", DEFAULT_GRID_BLOCKS, reset)
    
    'Section lines frequency
    txtSections(0).Text = LoadSettingInt("GridSectionsX", DEFAULT_GRID_SECTIONS, reset)
    txtSections(1).Text = LoadSettingInt("GridSectionsY", DEFAULT_GRID_SECTIONS, reset)

    'Grid offset
    GridOffset(0).value = LoadSettingInt("GridOffsetX", 0, reset)
    GridOffset(1).value = LoadSettingInt("GridOffsetY", 0, reset)
    
    chkSpaceY.value = IIf(txtBlocks(0).Text = txtBlocks(1).Text And _
                          txtSections(0).Text = txtSections(1).Text And _
                          GridOffset(0).value = GridOffset(1).value, _
                            vbChecked, vbUnchecked)
                            
    Call GridOffset_Change(0)
    Call GridOffset_Change(1)
    
    
    'Grid colors
    '0 Center: RGB(0, 0, 200)
    '1 Sections: RGB(200, 0, 0)
    '2 Blocks: RGB(150, 150, 150)
    '3 Std Grid: RGB(80, 80, 80)
    '(+4 for Y axis)
    
    lbl_color(0).BackColor = LoadSettingLng("GridColor0", DEFAULT_GRID_COLOR0, reset)
    lbl_color(1).BackColor = LoadSettingLng("GridColor1", DEFAULT_GRID_COLOR1, reset)
    lbl_color(2).BackColor = LoadSettingLng("GridColor2", DEFAULT_GRID_COLOR2, reset)
    lbl_color(3).BackColor = LoadSettingLng("GridColor3", DEFAULT_GRID_COLOR3, reset)
    lbl_color(4).BackColor = LoadSettingLng("GridColor4", DEFAULT_GRID_COLOR0, reset)
    lbl_color(5).BackColor = LoadSettingLng("GridColor5", DEFAULT_GRID_COLOR1, reset)
    lbl_color(6).BackColor = LoadSettingLng("GridColor6", DEFAULT_GRID_COLOR2, reset)
    lbl_color(7).BackColor = LoadSettingLng("GridColor7", DEFAULT_GRID_COLOR3, reset)
    
    chkColorY.value = vbChecked
    For i = 0 To 3
        If lbl_color(i).BackColor <> lbl_color(i + 4).BackColor Then
            chkColorY.value = vbUnchecked
        End If
    Next
    Call CheckColorY

    'Default tileset
    ret = LoadSetting("DefaultTileset", "", reset)
    If FileExists(ret) Then
        optDefaultTileset(1).value = True
        optDefaultTileset(0).value = False
        Call SetDefaultTileset(ret)
        tilesetpath = ret
    Else
        optDefaultTileset(0).value = True
        optDefaultTileset(1).value = False
        Call SetDefaultTileset("None")
    End If
    picdefaulttileset.Picture = LoadPicture(tilesetpath)
    picdefaulttileset.Refresh
    

    'Default walltiles
    ret = LoadSetting("DefaultWalltiles", "", reset)
    If FileExists(ret) Then
        optDefaultWall(1).value = True
        optDefaultWall(0).value = False
        wallsetpath = ret
        Call walltiles.LoadWallTiles(wallsetpath)
    Else
        optDefaultWall(0).value = True
        optDefaultWall(1).value = False
        wallsetpath = ""
    End If

    Call SetDefaultWall(wallsetpath, optDefaultWall(1).value)

    'Show tips at startup
    chkTips.value = LoadSettingInt("ShowTips", 1, reset)

    'Auto update (default set to "every day")
    chkAutoUpdate.value = LoadSettingInt("AutoUpdate", 1, reset)
    cmbUpdateDelay.ListIndex = LoadSettingInt("AutoUpdateDelay", 1, reset)

    cmbUpdateDelay.Enabled = (chkAutoUpdate.value = vbChecked)

    txtUpdateURL.Text = LoadSetting("UpdateURL", DEFAULT_UPDATE_URL, reset)
    
    chkUpdateURLCustom.value = IIf(txtUpdateURL.Text <> DEFAULT_UPDATE_URL, vbChecked, vbUnchecked)
    txtUpdateURL.Enabled = (txtUpdateURL.Text <> DEFAULT_UPDATE_URL)

    
    'New map on startup?
    chkNewMap.value = LoadSettingInt("AutoNewMap", 1, reset)




    'Display settings
    chkTilePreview.value = LoadSettingInt("ShowPreview", 1, reset)
    lbl_color(8).BackColor = LoadSettingLng("CursorColor", DEFAULT_CURSOR_COLOR, reset)
    chkShowTileCoord.value = LoadSettingInt("ShowCursorCoords", 1, reset)
        
    udRegnOpacity.value = LoadSettingInt("RegnOpacity1", DEFAULT_REGNOPACITY1, reset)
    udRegnOpacity2.value = LoadSettingInt("RegnOpacity2", DEFAULT_REGNOPACITY2, reset)
'    chkOnlyTopRegion.value = LoadSettingInt("RegnDrawTopOnly", 0, reset)

    
    lblLeftColor_color.BackColor = LoadSettingLng("LeftColor", DEFAULT_LEFTCOLOR, reset)
    lblRightColor_color.BackColor = LoadSettingLng("RightColor", DEFAULT_RIGHTCOLOR, reset)
    lblTilesetBackground_color.BackColor = LoadSettingLng("TilesetBackground", DEFAULT_TILESETBACKGROUND, reset)
    
    Call SetSetting("LeftColor", lblLeftColor_color.BackColor)
    Call SetSetting("RightColor", lblRightColor_color.BackColor)
    Call SetSetting("TilesetBackground", lblTilesetBackground_color.BackColor)
    
    
    
    'Autosave settings
    updownMaxAutosaves.value = LoadSettingInt("MaxAutosaves", DEFAULT_MAX_AUTOSAVES, reset)
    UpDownAutosave.value = LoadSettingInt("AutoSaveDelay", DEFAULT_AUTOSAVE_DELAY, reset)
    chkAutosave.value = LoadSettingInt("AutoSaveEnable", 1, reset)

    
    'Full map radar
    chkAutoFullMapPreview.value = LoadSettingInt("AutoFullMapPreview", 0, reset)
    
    
    
    chkLVZImagesAnimatedTiles.value = LoadSettingInt("AnimatedLVZImageTiles", 1, reset)
    


    'File associations
    If IsLVLAssociatedToDCME Then
        lblLVLAssociation.Caption = "LVL files are currently associated with DCME."
    Else
        lblLVLAssociation.Caption = "LVL files are associated with another program."
    End If
    
    ret = GetSystemImageEditor
    Call SetImageEditor(LoadSetting("ImageEditor", ret, reset))
    
    cmbImageEditor.list(1) = "Default: " & GetFileTitle(ret)
    
    
    Call UpdatePreview
End Sub


Private Sub DrawTilesetSettingsPreview()
    Const offsetX As Integer = 40
    Const offsetY As Integer = 6
    
    picPreviewTileset.Cls
    
    'Draw a portion of a tileset
    BitBlt picPreviewTileset.hDc, offsetX, offsetY, picPreviewTileset.width - offsetX, picPreviewTileset.Height - offsetY, frmGeneral.picdefaulttileset.hDc, 0, 0, vbSrcCopy
    
    shpPreviewTileset.BorderColor = lblLeftColor_color.BackColor
    
    Call DrawRectangle(picPreviewTileset.hDc, offsetX, offsetY, offsetX + 3 * TILEW, offsetY + 2 * TILEH, lblLeftColor_color.BackColor)
    Call DrawRectangle(picPreviewTileset.hDc, offsetX + TILEW, offsetY, offsetX + 2 * TILEW, offsetY + TILEH, lblRightColor_color.BackColor)
    
    'Draw the enlarged tile preview
    Call DrawImagePreviewCoords(frmGeneral.picdefaulttileset.hDc, 0, 0, 3 * TILEW, 2 * TILEH, picPreviewTileset.hDc, shpPreviewTileset.Left, shpPreviewTileset.Top, shpPreviewTileset.width, shpPreviewTileset.Height, lblTilesetBackground_color.BackColor)
    
    
    picPreviewTileset.Refresh
End Sub


Private Sub SetImageEditor(path As String, Optional Caption As String = "")
    
    If Not FileExists(path) Then path = GetSystemImageEditor
    
    picIconImageEditor.Cls
    
    If DrawFileIconOn(path, picIconImageEditor.hDc, 0, 0) Then
        If Caption <> "" Then
            cmbImageEditor.list(0) = Caption
        Else
            cmbImageEditor.list(0) = GetFileTitle(path)
        End If
        
        If cmbImageEditor.ListIndex <> 0 Then cmbImageEditor.ListIndex = 0
        
'        lblImageEditor.Caption = GetFileTitle(path)
        imageeditorpath = path
    End If
    
End Sub

Private Sub DrawWallTiles()
    Call walltiles.DrawWallTiles(picWalltiles.hDc, 4)
    

    picWalltiles.Refresh

    SetStretchBltMode picWallTilesPrev.hDc, HALFTONE
    StretchBlt picWallTilesPrev.hDc, 0, 0, picWallTilesPrev.ScaleWidth, picWallTilesPrev.ScaleHeight, picWalltiles.hDc, 0, 0, picWalltiles.width, picWalltiles.Height, vbSrcCopy

    picWallTilesPrev.Refresh
End Sub


Private Sub txtUpdateURL_LostFocus()
    If GetExtension(txtUpdateURL.Text) <> "txt" And GetExtension(txtUpdateURL.Text) <> "ini" Then
        MessageBox "Invalid URL. The URL provided must be a text file", vbOKOnly + vbExclamation
        txtUpdateURL.setfocus
        txtUpdateURL.selstart = Len(txtUpdateURL.Text)
    End If
End Sub



Private Sub UpDownAutosave_Change()
    chkAutosave.value = vbChecked
End Sub









Private Function LoadSetting(Key As String, defaultval As Variant, reset As Boolean) As String
    LoadSetting = IIf(reset, defaultval, GetSetting(Key, defaultval))
End Function

Private Function LoadSettingInt(Key As String, defaultval As Integer, reset As Boolean) As Integer
    LoadSettingInt = CInt(IIf(reset, defaultval, GetSetting(Key, defaultval)))
End Function

Private Function LoadSettingLng(Key As String, defaultval As Long, reset As Boolean) As Long
    LoadSettingLng = CLng(IIf(reset, defaultval, GetSetting(Key, defaultval)))
End Function
