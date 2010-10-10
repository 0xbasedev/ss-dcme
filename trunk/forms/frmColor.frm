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
    colorptr = ptr
    
    defaultColor = default
    oldcolor = old
    color = old
    
    filledpreview = filledprev

    cmd_default.visible = showdefault
    
    Call UpdatePointer
    Call UpdateValues
    
End Sub

Private Sub SetColor(val As Long)
    If colorptr Then
        CopyMemory ByVal colorptr, ByVal VarPtr(val), Len(val)
    Else
        
    End If
End Sub

Private Sub Cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_current_Click()

    color = oldcolor
    
    Call UpdatePointer
    Call UpdateValues

End Sub


Private Sub cmd_default_Click()

    color = defaultColor
    
    Call UpdatePointer
    Call UpdateValues

End Sub

Private Sub cmd_OK_Click()
    Call SetColor(color)
    
    Unload Me
End Sub

Private Sub colormap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X <= 0 Then X = 0
    If X >= colormap.ScaleWidth - 1 Then X = colormap.ScaleWidth - 1
    If Y <= 0 Then Y = 0
    If Y >= colormap.ScaleHeight - 1 Then Y = colormap.ScaleHeight - 1

    color = GetPixel(colormap.hDC, X, Y)
    Call UpdateValues
    Call PlacePointer(X, Y)

End Sub

Private Sub colormap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then Call colormap_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Form_Load()

    Set Me.Icon = frmGeneral.Icon

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub txtRGB_Change(Index As Integer)
    Call removeDisallowedCharacters(txtRGB(Index), 0, 255)
End Sub

Private Sub txtRGB_GotFocus(Index As Integer)
    txtRGB(Index).selstart = 0
    txtRGB(Index).sellength = Len(txtRGB(Index).Text)
End Sub

Private Sub txtRGB_LostFocus(Index As Integer)
    color = RGB(val(txtRGB(0).Text), val(txtRGB(1).Text), val(txtRGB(2).Text))

    Call UpdatePointer
    Call UpdateValues(Index)
End Sub

Private Sub UpdatePointer()
    Dim coord() As Single
    ReDim coord(1) As Single
    coord = FindColor(color)
    Call PlacePointer(coord(0), coord(1))
End Sub

Private Sub UpdateValues(Optional ignoreindex As Integer = -1)

    If ignoreindex <> 0 Then txtRGB(0).Text = GetRED(color)
    If ignoreindex <> 1 Then txtRGB(1).Text = GetGREEN(color)
    If ignoreindex <> 2 Then txtRGB(2).Text = GetBLUE(color)

    preview1.BorderColor = color
    preview2.BorderColor = color
    preview3.FillColor = color
    If filledpreview = True Then
        preview3.FillStyle = 0
        preview1.visible = False
        preview2.visible = False
    Else
        preview3.FillStyle = 1
        preview1.visible = True
        preview2.visible = True
    End If

End Sub

Private Sub PlacePointer(X As Single, Y As Single)
    colorpicker.Left = X - 3
    colorpicker.Top = Y - 3

    colorline1.x1 = X
    colorline1.x2 = X
    colorline1.y1 = Y - 5
    colorline1.y2 = Y + 5

    colorline2.x1 = X - 5
    colorline2.x2 = X + 5
    colorline2.y1 = Y
    colorline2.y2 = Y
End Sub

Private Function FindColor(color As Long) As Single()
    Dim retVal() As Single
    ReDim retVal(1) As Single
    Dim i As Integer
    Dim j As Integer

    Dim curpixel As Long
    Dim curdelta As Integer

    Dim mindelta As Integer
    Dim minx As Single
    Dim miny As Single

    mindelta = 255 * 3

    For j = 0 To colormap.ScaleWidth
        For i = 0 To colormap.ScaleHeight
            curpixel = GetPixel(colormap.hDC, i, j)
            curdelta = Abs(GetRED(curpixel) - GetRED(color)) + Abs(GetGREEN(curpixel) - GetGREEN(color)) + Abs(GetBLUE(curpixel) - GetBLUE(color))
            If curdelta < mindelta Then
                mindelta = curdelta
                minx = i
                miny = j
            End If
            If mindelta = 0 Then
                retVal(0) = minx
                retVal(1) = miny
                FindColor = retVal
                Exit Function
            End If
        Next
    Next
    retVal(0) = minx
    retVal(1) = miny
    FindColor = retVal
End Function

