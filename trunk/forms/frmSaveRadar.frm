VERSION 5.00
Begin VB.Form frmSaveRadar 
   Caption         =   "Save Screenshot"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fOptions 
      Caption         =   "Options"
      Height          =   1695
      Left            =   4680
      TabIndex        =   16
      Top             =   4200
      Width           =   2895
      Begin VB.CheckBox chkSelectionOnly 
         Caption         =   "Selection Only"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkDrawGrid 
         Caption         =   "Draw Grid"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkDrawLVZ 
         Caption         =   "Draw LVZ Objects"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fPreview 
      Caption         =   "Preview"
      Height          =   1215
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Frame fArea 
      Caption         =   "Area"
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3015
      Begin VB.PictureBox picradar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   600
         MousePointer    =   15  'Size All
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   10
         Top             =   480
         Width           =   1815
         Begin VB.Shape shparea 
            BorderColor     =   &H000000FF&
            Height          =   495
            Left            =   720
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.TextBox txtBottom 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "1023"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Text            =   "1023"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame fZoom 
      Caption         =   "Zoom"
      Height          =   1695
      Left            =   3360
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:16"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:4"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:2"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "2:1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin DCME.cPicViewer cPicPreview 
      Height          =   3615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7215
      _ExtentX        =   7223
      _ExtentY        =   5106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageWidth      =   65
      ImageHeight     =   57
      AnimationTime   =   0
   End
   Begin VB.Label lblSize 
      Caption         =   "Screenshot size: 16000 x 16000"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   6120
      Width           =   3015
   End
End
Attribute VB_Name = "frmSaveRadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Dim disp As New clsDisplayLayer

Dim Boundaries As area
Dim zoom As Single
Dim tileWidth As Integer, tileHeight As Integer
Dim fullwidth As Integer, fullheight As Integer



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub setParent(Main As frmMain)
    Set parent = Main
    
    If Not parent Is Nothing Then
    
        Call parent.cpic1024.stretchToDC(picradar.hDC, 0, 0, picradar.ScaleWidth, picradar.ScaleHeight, 0, 0, 1024, 1024, vbSrcCopy)
    
        zoom = parent.currentzoom
    
    End If
End Sub

Private Sub Form_Load()
    Boundaries.Left = 0
    Boundaries.Top = 0
    Boundaries.Right = MAPW - 1
    Boundaries.Bottom = MAPH - 1
    zoom = 1#

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
    disp.Cls
End Sub


Private Sub UpdateSize()
    tileWidth = (Boundaries.Right - Boundaries.Left + 1)
    tileHeight = (Boundaries.Bottom - Boundaries.Top + 1)
    
    fullwidth = tileWidth * TILEW * zoom
    fullheight = tileHeight * TILEH * zoom
    
    If fullwidth > 16383 Then fullwidth = 16383
    If fullheight > 16383 Then fullheight = 16383
    
    lblSize = "Screenshot Size: " & fullwidth & "x" & fullheight
    
    
    shparea.Left = (Boundaries.Left / MAPW) * picradar.ScaleWidth
    shparea.Top = (Boundaries.Top / MAPH) * picradar.ScaleHeight
    shparea.width = (tileWidth / MAPW) * picradar.ScaleWidth
    shparea.height = (tileHeight / MAPH) * picradar.ScaleHeight
    
    
    
    Call Render
End Sub


Private Sub Render()
    disp.BackColor = vbBlack
    Call disp.Resize(fullwidth, fullheight, False)
    
    
    
    cPicPreview.width = cPicPreview.width
    cPicPreview.height = cPicPreview.height
    cPicPreview.BackColor = vbBlack

    cPicPreview.imageWidth = fullwidth
    cPicPreview.imageHeight = fullheight
    DrawFilledRectangle cPicPreview.hDC, 0, 0, fullwidth, fullheight, vbBlack
    Call parent.RenderTiles(Boundaries.Left, Boundaries.Top, 0, 0, fullwidth, fullheight, cPicPreview.hDC, False, False)
    
'    Call disp.SaveToFile("C:\Jeux\Continuum\DCME\bleh.bmp", False)
    Call cPicPreview.Refresh

    
'    cPicPreview.Clear
    
'    BitBlt Picture1.hDC, 0, 0, Picture1.width, Picture1.height, disp.hDC, 0, 0, vbSrcCopy
    
'    BitBlt cPicPreview.hDC, 0, 0, fullWidth, fullHeight, disp.hDC, 0, 0, vbSrcCopy

End Sub



Private Sub txtBottom_Change()
    Call removeDisallowedCharacters(txtBottom, 0, 1023, False)
    Boundaries.Bottom = CInt(val(txtBottom.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub

Private Sub txtLeft_Change()
    Call removeDisallowedCharacters(txtLeft, 0, 1023, False)
    Boundaries.Left = CInt(val(txtLeft.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub

Private Sub txtRight_Change()
    Call removeDisallowedCharacters(txtRight, 0, 1023, False)
    Boundaries.Right = CInt(val(txtRight.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub

Private Sub txtTop_Change()
    Call removeDisallowedCharacters(txtTop, 0, 1023, False)
    Boundaries.Top = CInt(val(txtTop.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub
