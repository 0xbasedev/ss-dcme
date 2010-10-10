VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jump to..."
   ClientHeight    =   1740
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.PictureBox picmap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.Line centerYLine 
         BorderColor     =   &H00FF0000&
         X1              =   52
         X2              =   52
         Y1              =   0
         Y2              =   105
      End
      Begin VB.Line centerXLine 
         BorderColor     =   &H00FF0000&
         X1              =   0
         X2              =   105
         Y1              =   52
         Y2              =   52
      End
      Begin VB.Line cursorLine1 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         X1              =   32
         X2              =   44
         Y1              =   56
         Y2              =   56
      End
      Begin VB.Line cursorLine2 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         X1              =   32
         X2              =   32
         Y1              =   56
         Y2              =   48
      End
      Begin VB.Shape cursor 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         Height          =   105
         Left            =   720
         Shape           =   3  'Circle
         Top             =   480
         Width           =   105
      End
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "Jump"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txt_Y 
      Height          =   285
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txt_X 
      Height          =   285
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   8
      Y2              =   112
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Xcoord As Integer
Public Ycoord As Integer

Dim ignoreupdate As Boolean

Private Sub cmd_OK_Click()
10        txt_X_LostFocus
20        txt_Y_LostFocus

30        If IsNumeric(txt_Y.Text) And val(txt_Y.Text) >= 0 And val(txt_Y.Text) <= 1023 _
             And IsNumeric(txt_X.Text) And val(txt_X.Text) >= 0 And val(txt_X.Text) <= 1023 Then
40            Call frmGeneral.ExecuteGoTo(Xcoord, Ycoord)
50            Unload Me
60        End If
End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon

20        If Xcoord > 1023 Then Xcoord = 1023
30        If Xcoord < 0 Then Xcoord = 0
40        If Ycoord > 1023 Then Ycoord = 1023
50        If Ycoord < 0 Then Ycoord = 0

60        txt_X.Text = Xcoord
70        txt_Y.Text = Ycoord
80        Call SetPointerToValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        SharedVar.MouseDown = 0
End Sub

Private Sub picmap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        SharedVar.MouseDown = Button

20        If X <= 0 Then X = 0
30        If X >= picmap.ScaleWidth - 1 Then X = picmap.ScaleWidth - 1
40        If Y <= 0 Then Y = 0
50        If Y >= picmap.ScaleHeight - 1 Then Y = picmap.ScaleHeight - 1

60        Call PlacePointer(X, Y)
70        Call UpdateValues
End Sub

Private Sub picmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If SharedVar.MouseDown <> 0 Then
20            Call picmap_MouseDown(Button, Shift, X, Y)
30        End If

End Sub

Private Sub picmap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        SharedVar.MouseDown = 0
End Sub

Private Sub txt_X_Change()
10        Call removeDisallowedCharacters(txt_X, 0, 1024)

20        If Len(txt_X.Text) = 4 Then
30            txt_Y.setfocus
40        End If

50        Call SetPointerToValues
End Sub

Private Sub txt_Y_Change()
10        Call removeDisallowedCharacters(txt_Y, 0, 1024)

20        If Len(txt_Y.Text) = 4 Then
30            txt_X.setfocus
40        End If

50        Call SetPointerToValues
End Sub
Private Sub txt_X_GotFocus()
10        txt_X.selstart = 0
20        txt_X.sellength = Len(txt_X.Text)
End Sub

Private Sub txt_X_LostFocus()
10        Xcoord = check_coord(txt_X.Text)
20        txt_X.Text = Xcoord
End Sub

Private Sub txt_Y_GotFocus()
10        txt_Y.selstart = 0
20        txt_Y.sellength = Len(txt_Y.Text)
End Sub

Private Sub txt_Y_LostFocus()
10        Ycoord = check_coord(txt_Y.Text)
20        txt_Y.Text = Ycoord
End Sub

Private Function check_coord(coord As String) As Integer
10        If IsNumeric(coord) Then
20            If coord > 1023 Then coord = 1023
30            If coord < 0 Then coord = 0
40            check_coord = Int(coord)
50        Else
60            check_coord = 512
70        End If
End Function

Private Sub PlacePointer(X As Single, Y As Single)
10        cursor.Left = X - 3
20        cursor.Top = Y - 3

30        cursorLine1.x1 = X
40        cursorLine1.x2 = X
50        cursorLine1.y1 = Y - 5
60        cursorLine1.y2 = Y + 5

70        cursorLine2.x1 = X - 5
80        cursorLine2.x2 = X + 5
90        cursorLine2.y1 = Y
100       cursorLine2.y2 = Y
End Sub

Private Sub SetPointerToValues()
10        If ignoreupdate Then Exit Sub

          Dim X As Single
          Dim Y As Single
20        X = Int((val(txt_X.Text) / 1024) * (picmap.width - 1))
30        Y = Int((val(txt_Y.Text) / 1024) * (picmap.height - 1))

40        Call PlacePointer(X, Y)
End Sub

Private Sub UpdateValues()
      ' 0.. picmap.width = Int(cursor.Left + cursor.Width / 2)
10        ignoreupdate = True
20        txt_X.Text = Int(((cursor.Left + 3) / (picmap.width - 1)) * 1023)
30        txt_Y.Text = Int(((cursor.Top + 3) / (picmap.height - 1)) * 1023)
40        ignoreupdate = False

End Sub

