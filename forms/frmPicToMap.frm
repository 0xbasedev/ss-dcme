VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPicToMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picture to Map"
   ClientHeight    =   3990
   ClientLeft      =   345
   ClientTop       =   615
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkInverted 
      Caption         =   "Inverted pixel check (black = tile, else no tile)"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   195
      Width           =   1095
   End
   Begin VB.PictureBox picpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   4440
      ScaleHeight     =   3495
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picpreview 
      AutoRedraw      =   -1  'True
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lblname 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmPicToMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'holds the tile that is used, this will be set when the form is loaded
'from the general form
Private parent As frmMain


Public Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Private Sub cmdBrowse_Click()
      'Opens a dialog for selecting a picture (bmp or jpg)

      'dialog settings
10        On Error GoTo errh
20        cd.DialogTitle = "Open a picture..."
30        cd.flags = cdlOFNHideReadOnly
40        cd.Filter = "Supported image files (*.bmp, *.png, *.jpg, *.gif)|*.bmp;*.jpg;*.png;*.gif;*.bm2;*.jpeg"
50        cd.ShowOpen

          'update the label with the filename
60        lblname.Caption = cd.filetitle

          'load the picture in the picturebox and give it the correct size
70        Call LoadPic(picpic, cd.filename)
80        picpic.AutoSize = True
90        picpic.AutoSize = False

          'make sure the picture isn't bigger than 1024 pixels, or it would be bigger
          'than the map itself
100       If picpic.width > 1024 Then
110           picpic.width = 1024
120       End If

130       If picpic.height > 1024 Then
140           picpic.height = 1024
150       End If

          'show the preview
160       SetStretchBltMode picpreview.hDC, HALFTONE
170       StretchBlt picpreview.hDC, 0, 0, picpreview.width, picpreview.height, picpic.hDC, 0, 0, picpic.width, picpic.height, vbSrcCopy
180       picpreview.Refresh

          'blt the picture onto the image of the picture box
190       BitBlt picpic.hDC, 0, 0, picpic.width, picpic.height, picpic.hDC, 0, 0, vbSrcCopy
200       picpic.Refresh

          'preparations are now ready, enable the go
210       cmdGo.Enabled = True

220       Exit Sub
errh:
          'pressed cancel in dialog box, do nothing
230       If Err = cdlCancel Then
240           Exit Sub
250       End If

End Sub

Private Sub cmdCancel_Click()
      'Cancels the form
      'enable the general form again and unload this form
10        Set parent = Nothing
          
20        Unload Me
End Sub

Private Sub cmdGo_Click()
      'Do the diffusion, build the array and pass it to the general form, afterwards unload this form
      ' use error diffusion dithering to convert the image to black and
      ' white, then scan like TTM
10        On Error GoTo cmdGo_Click_Error
          
20        frmGeneral.IsBusy("frmPicToMap.cmgGo_Click") = True
          
          'Create a monochrome pic
30        Call FloydSteinberg(picpic)

          'now create an array
          Dim Pic() As Integer
40        ReDim Pic(picpic.width, picpic.height)

          Dim i As Integer, j As Integer
          Dim TileToUse As Integer
          
50        If parent.tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
60            TileToUse = parent.tileset.selection(vbLeftButton).tilenr
70        Else
80            TileToUse = 1
90        End If
          
          Dim isInverted As Boolean
100       isInverted = (chkInverted.value = vbChecked)
          
          Dim c As Long
          
          'check every pixel and see if it's black or not
110       For j = 0 To picpic.height
120           For i = 0 To picpic.width
130               c = GetPixel(picpic.hDC, i, j)

                  'if option inverted is on then switch the statements
140               If (c = vbBlack) = isInverted Then
150                   Pic(i, j) = TileToUse
160               Else
170                   Pic(i, j) = 0
180               End If
190           Next
200       Next

          'reenable the general form
          'pass the info to the ExecutePicToMap and unload this form
          'Call frmGeneral.ExecutePicToMap(Pic, picpic.width, picpic.height)
210       Call parent.sel.PicToMap(Pic, picpic.width, picpic.height)
          
220       frmGeneral.IsBusy("frmPicToMap.cmgGo_Click") = False
          
230       Unload Me

          
240       On Error GoTo 0
250       Exit Sub

cmdGo_Click_Error:
260       frmGeneral.IsBusy("frmPicToMap.cmgGo_Click") = False
270       HandleError Err, "frmPicToMap.cmgGo_Click"
End Sub

Private Sub Form_Load()
      'Disable the general form and go button
10        Set Me.Icon = frmGeneral.Icon
20        cmdGo.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      'If we press X, then do cancel
10        cmdCancel_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
End Sub
