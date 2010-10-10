VERSION 5.00
Begin VB.Form frmSaveRadar 
   Caption         =   "Save Radar View"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox piclevel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15360
      Left            =   8640
      ScaleHeight     =   1024
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   64
      Left            =   0
      SmallChange     =   8
      TabIndex        =   2
      Top             =   3600
      Width           =   3615
   End
   Begin VB.VScrollBar VScroll 
      Height          =   2895
      LargeChange     =   64
      Left            =   3720
      SmallChange     =   8
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtDimension 
      Height          =   285
      Left            =   6240
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "1024"
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      MousePointer    =   15  'Size All
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblDimension 
      Alignment       =   2  'Center
      Caption         =   "Size :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "frmSaveRadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dragX As Integer
Dim dragY As Integer

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub cmdSave_Click()
10        If val(txtDimension.Text) <> 0 Then
20            Call frmGeneral.SaveMiniMap("", val(txtDimension.Text))
30            Unload Me
40        Else
50            txtDimension.setfocus
60            txtDimension.selstart = 1
70            txtDimension.sellength = Len(txtDimension.Text)
80        End If
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon

20        Me.width = (520 + VScroll.width) * Screen.TwipsPerPixelX
30        Me.height = (526 + txtDimension.height + HScroll.height + cmdSave.height) * Screen.TwipsPerPixelY

40        Call UpdatePreview

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Unload Me
End Sub

Private Sub HScroll_Change()
10        Call UpdatePreview

20        picpreview.setfocus
End Sub



Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        dragX = X
20        dragY = Y
End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim tmp As Integer

10        If Button Then
20            If HScroll.Enabled Then
30                tmp = HScroll.value + (dragX - X)
40                If tmp >= HScroll.Max Then
50                    HScroll.value = HScroll.Max
60                ElseIf tmp <= 0 Then
70                    HScroll.value = 0
80                Else
90                    HScroll.value = tmp
100               End If
110           End If

120           If VScroll.Enabled Then
130               tmp = VScroll.value + (dragY - Y)
140               If tmp >= VScroll.Max Then
150                   VScroll.value = VScroll.Max
160               ElseIf tmp <= 0 Then
170                   VScroll.value = 0
180               Else
190                   VScroll.value = tmp
200               End If
210           End If

220           Call UpdatePreview

230           dragX = X
240           dragY = Y
250       End If
End Sub

Private Sub VScroll_Change()
10        Call UpdatePreview

20        picpreview.setfocus
End Sub

Private Sub txtDimension_Change()
10        Call removeDisallowedCharacters(txtDimension, 0, 1024)

20        If val(txtDimension) <> 0 Then
30            Call UpdatePreview
40        End If

End Sub

Private Sub UpdatePreview()
          Dim dimension As Integer
          Dim previewsizeX As Integer
          Dim previewsizeY As Integer

10        dimension = val(txtDimension.Text)

20        previewsizeX = Me.ScaleWidth - VScroll.width
30        previewsizeY = Me.ScaleHeight - txtDimension.height - HScroll.height - cmdSave.height - 6


40        If dimension > previewsizeX Then
50            picpreview.Left = 0
60            picpreview.width = previewsizeX

70            HScroll.Enabled = True
80            HScroll.Max = dimension - picpreview.width
90        Else
100           picpreview.width = dimension
110           picpreview.Left = (previewsizeX - dimension) \ 2

120           HScroll.Enabled = False
130           HScroll.value = 0
140       End If



150       If dimension > previewsizeY Then
160           picpreview.Top = 0
170           picpreview.height = previewsizeY

180           VScroll.Enabled = True
190           VScroll.Max = dimension - picpreview.height
200       Else
210           picpreview.height = dimension
220           picpreview.Top = (previewsizeY - dimension) \ 2

230           VScroll.Enabled = False
240           VScroll.value = 0
250       End If

260       picpreview.Cls


270       If dimension = 1024 Then
              'Same size, use bitblt
280           BitBlt picpreview.hDC, 0, 0, picpreview.width, picpreview.height, piclevel.hDC, HScroll.value, VScroll.value, vbSrcCopy
290       Else
              'Source is smaller, use halftone resize
300           SetStretchBltMode picpreview.hDC, HALFTONE
310           StretchBlt picpreview.hDC, -HScroll.value, -VScroll.value, dimension, dimension, piclevel.hDC, 0, 0, piclevel.width, piclevel.height, vbSrcCopy
320       End If


330       picpreview.Refresh



End Sub


Private Sub Form_Resize()

          Dim previewsizeX As Integer
          Dim previewsizeY As Integer


10        If Me.height < (txtDimension.height + HScroll.height + cmdSave.height + 134) * Screen.TwipsPerPixelY Then
20            Me.height = (txtDimension.height + HScroll.height + cmdSave.height + 134) * Screen.TwipsPerPixelY
30        End If
40        If Me.width < (VScroll.width + 128) * Screen.TwipsPerPixelX Then
50            Me.width = (VScroll.width + 128) * Screen.TwipsPerPixelX
60        End If

70        previewsizeX = Me.ScaleWidth - VScroll.width
80        previewsizeY = Me.ScaleHeight - txtDimension.height - HScroll.height - cmdSave.height - 4

90        HScroll.width = previewsizeX
100       HScroll.Top = previewsizeY
110       VScroll.Left = previewsizeX
120       VScroll.height = previewsizeY

130       lblDimension.Top = previewsizeY + HScroll.height

140       txtDimension.Left = lblDimension.width + 5
150       txtDimension.Top = previewsizeY + HScroll.height
160       txtDimension.width = Me.ScaleWidth - lblDimension.width

170       cmdSave.width = Me.ScaleWidth \ 2
180       cmdCancel.width = Me.ScaleWidth \ 2
190       cmdCancel.Left = cmdSave.width
200       cmdSave.Top = previewsizeY + lblDimension.height + HScroll.height + 2
210       cmdCancel.Top = cmdSave.Top





220       Call UpdatePreview

End Sub


