VERSION 5.00
Begin VB.Form frmSaveRadar 
   Caption         =   "Save Screenshot"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   223
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fOptions 
      Caption         =   "Options"
      Height          =   1455
      Left            =   1200
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CheckBox chkSelectionOnly 
         Caption         =   "Selection Only"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkDrawGrid 
         Caption         =   "Draw Grid"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkDrawLVZ 
         Caption         =   "Draw LVZ Objects"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fArea 
      Caption         =   "Area"
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3255
      Begin VB.PictureBox picradar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   120
         ScaleHeight     =   201
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   9
         Top             =   480
         Width           =   3015
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
         Left            =   2640
         TabIndex        =   8
         Text            =   "520"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Text            =   "520"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "500"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "500"
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame fZoom 
      Caption         =   "Zoom"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:16"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:4"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:2"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optZoomlevel 
         Caption         =   "1:1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.PictureBox picfull 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   4200
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label lblSize 
      Caption         =   "Screenshot size: 16000 x 16000"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5760
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

Dim Boundaries As area
Dim tileWidth As Integer, tileHeight As Integer
Dim fullwidth As Integer, fullheight As Integer


Dim bFreezeTextboxes As Boolean


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub setParent(Main As frmMain)
    Set parent = Main
    
    If Not parent Is Nothing Then
        SetStretchBltMode picradar.hDC, HALFTONE
        Call parent.cpic1024.stretchToDC(picradar.hDC, 0, 0, picradar.ScaleWidth, picradar.ScaleHeight, 0, 0, 1024, 1024, vbSrcCopy)
        
        Boundaries.Left = parent.ScreenToTileX(0)
        Boundaries.Right = parent.ScreenToTileX(parent.picPreview.width - 1)
        Boundaries.Top = parent.ScreenToTileY(0)
        Boundaries.Bottom = parent.ScreenToTileY(parent.picPreview.height - 1)

        
        bFreezeTextboxes = True
        Call UpdateBoundaries
        bFreezeTextboxes = False
        Call UpdateSize
    End If
End Sub



Private Sub cmdSave_Click()
    
    If FileExists("screenshot.bmp") Then
        If Not CheckOverwrite("screenshot.bmp") Then
            Exit Sub
        End If
    End If

    Call Render
    
    Call SavePicture(picfull.Image, "screenshot.bmp")
End Sub

Private Sub Form_Load()
    bFreezeTextboxes = False
    Boundaries.Left = 0
    Boundaries.Right = MAPW
    Boundaries.Top = 0
    Boundaries.Bottom = MAPH
End Sub

Private Sub UpdateBoundaries()
    txtLeft.Text = CStr(Boundaries.Left)
    txtRight.Text = CStr(Boundaries.Right)
    txtTop.Text = CStr(Boundaries.Top)
    txtBottom.Text = CStr(Boundaries.Bottom)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
    picfull.width = 1
    picfull.height = 1
    picfull.Cls
    
End Sub


Private Sub UpdateSize()
    If parent Is Nothing Then Exit Sub
    
    tileWidth = (Boundaries.Right - Boundaries.Left + 1)
    tileHeight = (Boundaries.Bottom - Boundaries.Top + 1)
    
    fullwidth = tileWidth * parent.currenttilew
    fullheight = tileHeight * parent.currenttilew
    
    If fullwidth > 16380 Then fullwidth = 16380
    If fullheight > 16380 Then fullheight = 16380
    
    lblSize = "Screenshot Size: " & fullwidth & "x" & fullheight
    
    
    shparea.Left = (Boundaries.Left / MAPW) * picradar.ScaleWidth
    shparea.Top = (Boundaries.Top / MAPH) * picradar.ScaleHeight
    shparea.width = (tileWidth / MAPW) * picradar.ScaleWidth
    shparea.height = (tileHeight / MAPH) * picradar.ScaleHeight
    
    
    
    Call Render
End Sub


Private Sub Render()
 On Error GoTo render_error
 
'    disp.BackColor = vbBlack
'    Call disp.Resize(fullwidth, fullheight, False)
'
    picfull.width = fullwidth
    picfull.height = fullheight
    
'
'    cPicPreview.AnimFramesX = 1
'    cPicPreview.AnimFramesX = 1
'
'    cPicPreview.width = cPicPreview.width
'    cPicPreview.height = cPicPreview.height
'    cPicPreview.BackColor = vbBlack
'
'    cPicPreview.imageWidth = fullwidth
'    cPicPreview.imageHeight = fullheight
'    cPicPreview.Clear
    picfull.Cls
    
    'DrawFilledRectangle cPicPreview.hDC, 0, 0, fullwidth, fullheight, vbBlack
    Call parent.RenderTiles(Boundaries.Left, Boundaries.Top, 0, 0, fullwidth - 1, fullheight - 1, picfull.hDC, False, False)

    
    'Call cPicPreview.Refresh

    
'    cPicPreview.Clear
    picfull.Refresh
    
'    BitBlt cPicPreview.hDC, 0, 0, fullWidth, fullHeight, disp.hDC, 0, 0, vbSrcCopy
    Exit Sub
render_error:
    MsgBox "Cannot create image, try rendering a smaller image", vbExclamation
End Sub





Private Sub optZoomlevel_Click(Index As Integer)
    Dim zoom As Single
    
    If Not parent Is Nothing Then
        Select Case Index
        Case 0
            zoom = 2#
        Case 1
            zoom = 1#
        Case 2
            zoom = 0.5
        Case 3
            zoom = 0.25
        Case 4
            zoom = 1 / 16
        Case Else
            zoom = 1#
        End Select
        
        
        Call parent.magnifier.SetZoom(zoom, True)
        Call parent.SetFocusAt(Boundaries.Left + (Boundaries.Right - Boundaries.Left) \ 2, _
                                Boundaries.Top + (Boundaries.Bottom - Boundaries.Top) \ 2, _
                                parent.picPreview.width \ 2, parent.picPreview.height \ 2, True)
    End If
    Call UpdateSize
    
End Sub



Private Sub picradar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If inside the rectangle, show 'move' arrow
    
    'If on the edge, show 'resize' arrow


End Sub

Private Sub txtBottom_Change()
    If bFreezeTextboxes Then Exit Sub
    
    Call removeDisallowedCharacters(txtBottom, 0, 1023, False)
    Boundaries.Bottom = CInt(val(txtBottom.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub

Private Sub txtLeft_Change()
    If bFreezeTextboxes Then Exit Sub
    
    Call removeDisallowedCharacters(txtLeft, 0, 1023, False)
    Boundaries.Left = CInt(val(txtLeft.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub

Private Sub txtRight_Change()
    If bFreezeTextboxes Then Exit Sub
    
    Call removeDisallowedCharacters(txtRight, 0, 1023, False)
    Boundaries.Right = CInt(val(txtRight.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub

Private Sub txtTop_Change()
    If bFreezeTextboxes Then Exit Sub
    
    Call removeDisallowedCharacters(txtTop, 0, 1023, False)
    Boundaries.Top = CInt(val(txtTop.Text))
    
    If Boundaries.Right >= Boundaries.Left And Boundaries.Bottom >= Boundaries.Top Then Call UpdateSize
End Sub
