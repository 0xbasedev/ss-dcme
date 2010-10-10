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
    txt_X_LostFocus
    txt_Y_LostFocus

    If IsNumeric(txt_Y.Text) And val(txt_Y.Text) >= 0 And val(txt_Y.Text) <= 1023 _
       And IsNumeric(txt_X.Text) And val(txt_X.Text) >= 0 And val(txt_X.Text) <= 1023 Then
        Call frmGeneral.ExecuteGoTo(Xcoord, Ycoord)
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    If Xcoord > 1023 Then Xcoord = 1023
    If Xcoord < 0 Then Xcoord = 0
    If Ycoord > 1023 Then Ycoord = 1023
    If Ycoord < 0 Then Ycoord = 0

    txt_X.Text = Xcoord
    txt_Y.Text = Ycoord
    Call SetPointerToValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SharedVar.MouseDown = 0
End Sub

Private Sub picmap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SharedVar.MouseDown = Button

    If X <= 0 Then X = 0
    If X >= picmap.ScaleWidth - 1 Then X = picmap.ScaleWidth - 1
    If Y <= 0 Then Y = 0
    If Y >= picmap.ScaleHeight - 1 Then Y = picmap.ScaleHeight - 1

    Call PlacePointer(X, Y)
    Call UpdateValues
End Sub

Private Sub picmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SharedVar.MouseDown <> 0 Then
        Call picmap_MouseDown(Button, Shift, X, Y)
    End If

End Sub

Private Sub picmap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SharedVar.MouseDown = 0
End Sub

Private Sub txt_X_Change()
    Call removeDisallowedCharacters(txt_X, 0, 1024)

    If Len(txt_X.Text) = 4 Then
        txt_Y.setfocus
    End If

    Call SetPointerToValues
End Sub

Private Sub txt_Y_Change()
    Call removeDisallowedCharacters(txt_Y, 0, 1024)

    If Len(txt_Y.Text) = 4 Then
        txt_X.setfocus
    End If

    Call SetPointerToValues
End Sub
Private Sub txt_X_GotFocus()
    txt_X.selstart = 0
    txt_X.sellength = Len(txt_X.Text)
End Sub

Private Sub txt_X_LostFocus()
    Xcoord = check_coord(txt_X.Text)
    txt_X.Text = Xcoord
End Sub

Private Sub txt_Y_GotFocus()
    txt_Y.selstart = 0
    txt_Y.sellength = Len(txt_Y.Text)
End Sub

Private Sub txt_Y_LostFocus()
    Ycoord = check_coord(txt_Y.Text)
    txt_Y.Text = Ycoord
End Sub

Private Function check_coord(coord As String) As Integer
    If IsNumeric(coord) Then
        If coord > 1023 Then coord = 1023
        If coord < 0 Then coord = 0
        check_coord = Int(coord)
    Else
        check_coord = 512
    End If
End Function

Private Sub PlacePointer(X As Single, Y As Single)
    cursor.Left = X - 3
    cursor.Top = Y - 3

    cursorLine1.x1 = X
    cursorLine1.x2 = X
    cursorLine1.y1 = Y - 5
    cursorLine1.y2 = Y + 5

    cursorLine2.x1 = X - 5
    cursorLine2.x2 = X + 5
    cursorLine2.y1 = Y
    cursorLine2.y2 = Y
End Sub

Private Sub SetPointerToValues()
    If ignoreupdate Then Exit Sub

    Dim X As Single
    Dim Y As Single
    X = Int((val(txt_X.Text) / 1024) * (picmap.width - 1))
    Y = Int((val(txt_Y.Text) / 1024) * (picmap.height - 1))

    Call PlacePointer(X, Y)
End Sub

Private Sub UpdateValues()
' 0.. picmap.width = Int(cursor.Left + cursor.Width / 2)
    ignoreupdate = True
    txt_X.Text = Int(((cursor.Left + 3) / (picmap.width - 1)) * 1023)
    txt_Y.Text = Int(((cursor.Top + 3) / (picmap.height - 1)) * 1023)
    ignoreupdate = False

End Sub

