VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color map"
   ClientHeight    =   3195
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Choose a color"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtRGB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   840
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "B"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   840
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "G"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "R"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton cmd_cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmd_current 
         Caption         =   "Current"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmd_default 
         Caption         =   "Default"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmd_OK 
         Caption         =   "OK"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
      Begin VB.PictureBox colormap 
         AutoRedraw      =   -1  'True
         Height          =   1935
         Left            =   120
         MousePointer    =   2  'Cross
         Picture         =   "frmColor.frx":0000
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   117
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         Begin VB.Shape colorpicker 
            DrawMode        =   6  'Mask Pen Not
            Height          =   105
            Left            =   750
            Shape           =   3  'Circle
            Top             =   600
            Width           =   105
         End
         Begin VB.Line colorline2 
            DrawMode        =   6  'Mask Pen Not
            X1              =   40
            X2              =   40
            Y1              =   64
            Y2              =   56
         End
         Begin VB.Line colorline1 
            DrawMode        =   6  'Mask Pen Not
            X1              =   45
            X2              =   55
            Y1              =   64
            Y2              =   64
         End
      End
      Begin VB.Label lblB 
         Caption         =   "Blue:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblG 
         Caption         =   "Green:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblR 
         Caption         =   "Red:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lbl_preview 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Line preview2 
         BorderColor     =   &H000000FF&
         X1              =   2520
         X2              =   2520
         Y1              =   960
         Y2              =   480
      End
      Begin VB.Line preview1 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   2880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Shape preview3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   495
         Left            =   2160
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim color As Long

'Stores color that is currently set
Dim colorptr As Long
Dim Index As Integer

Dim defaultColor As Long
Dim oldcolor As Long

Dim filledpreview As Boolean

Sub SetData(old As Long, ByVal ptr As Long, default As Long, showdefault As Boolean, filledprev As Boolean)
10        colorptr = ptr
          
20        defaultColor = default
30        oldcolor = old
40        color = old
          
50        filledpreview = filledprev

60        cmd_default.visible = showdefault
          
70        Call UpdatePointer
80        Call UpdateValues
          
End Sub

Private Sub SetColor(val As Long)
10        If colorptr Then
20            CopyMemory ByVal colorptr, ByVal VarPtr(val), Len(val)
30        Else
              
40        End If
End Sub

Private Sub Cmd_cancel_Click()
10        Unload Me
End Sub

Private Sub cmd_current_Click()

10        color = oldcolor
          
20        Call UpdatePointer
30        Call UpdateValues

End Sub


Private Sub cmd_default_Click()

10        color = defaultColor
          
20        Call UpdatePointer
30        Call UpdateValues

End Sub

Private Sub cmd_OK_Click()
10        Call SetColor(color)
          
20        Unload Me
End Sub

Private Sub colormap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        If X <= 0 Then X = 0
20        If X >= colormap.ScaleWidth - 1 Then X = colormap.ScaleWidth - 1
30        If Y <= 0 Then Y = 0
40        If Y >= colormap.ScaleHeight - 1 Then Y = colormap.ScaleHeight - 1

50        color = GetPixel(colormap.hDC, X, Y)
60        Call UpdateValues
70        Call PlacePointer(X, Y)

End Sub

Private Sub colormap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If Button Then Call colormap_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Form_Load()

10        Set Me.Icon = frmGeneral.Icon

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Unload Me
End Sub

Private Sub txtRGB_Change(Index As Integer)
10        Call removeDisallowedCharacters(txtRGB(Index), 0, 255)
End Sub

Private Sub txtRGB_GotFocus(Index As Integer)
10        txtRGB(Index).selstart = 0
20        txtRGB(Index).sellength = Len(txtRGB(Index).Text)
End Sub

Private Sub txtRGB_LostFocus(Index As Integer)
10        color = RGB(val(txtRGB(0).Text), val(txtRGB(1).Text), val(txtRGB(2).Text))

20        Call UpdatePointer
30        Call UpdateValues(Index)
End Sub

Private Sub UpdatePointer()
          Dim coord() As Single
10        ReDim coord(1) As Single
20        coord = FindColor(color)
30        Call PlacePointer(coord(0), coord(1))
End Sub

Private Sub UpdateValues(Optional ignoreindex As Integer = -1)

10        If ignoreindex <> 0 Then txtRGB(0).Text = GetRED(color)
20        If ignoreindex <> 1 Then txtRGB(1).Text = GetGREEN(color)
30        If ignoreindex <> 2 Then txtRGB(2).Text = GetBLUE(color)

40        preview1.BorderColor = color
50        preview2.BorderColor = color
60        preview3.FillColor = color
70        If filledpreview = True Then
80            preview3.FillStyle = 0
90            preview1.visible = False
100           preview2.visible = False
110       Else
120           preview3.FillStyle = 1
130           preview1.visible = True
140           preview2.visible = True
150       End If

End Sub

Private Sub PlacePointer(X As Single, Y As Single)
10        colorpicker.Left = X - 3
20        colorpicker.Top = Y - 3

30        colorline1.x1 = X
40        colorline1.x2 = X
50        colorline1.y1 = Y - 5
60        colorline1.y2 = Y + 5

70        colorline2.x1 = X - 5
80        colorline2.x2 = X + 5
90        colorline2.y1 = Y
100       colorline2.y2 = Y
End Sub

Private Function FindColor(color As Long) As Single()
          Dim retVal() As Single
10        ReDim retVal(1) As Single
          Dim i As Integer
          Dim j As Integer

          Dim curpixel As Long
          Dim curdelta As Integer

          Dim mindelta As Integer
          Dim minx As Single
          Dim miny As Single

20        mindelta = 255 * 3

30        For j = 0 To colormap.ScaleWidth
40            For i = 0 To colormap.ScaleHeight
50                curpixel = GetPixel(colormap.hDC, i, j)
60                curdelta = Abs(GetRED(curpixel) - GetRED(color)) + Abs(GetGREEN(curpixel) - GetGREEN(color)) + Abs(GetBLUE(curpixel) - GetBLUE(color))
70                If curdelta < mindelta Then
80                    mindelta = curdelta
90                    minx = i
100                   miny = j
110               End If
120               If mindelta = 0 Then
130                   retVal(0) = minx
140                   retVal(1) = miny
150                   FindColor = retVal
160                   Exit Function
170               End If
180           Next
190       Next
200       retVal(0) = minx
210       retVal(1) = miny
220       FindColor = retVal
End Function

