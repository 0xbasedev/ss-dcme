VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerClose 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   6480
      Top             =   960
   End
   Begin VB.PictureBox picThanks 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   4080
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   6
      Top             =   2340
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   6480
      Top             =   480
   End
   Begin VB.TextBox txtThanks 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmSplash.frx":5289E
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblEmailSamapico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Samapico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5520
      MousePointer    =   2  'Cross
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblEmailDrake 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drake7707"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5520
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbllink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SSForum.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblvisitssforum 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit                         for info and updates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label lbladditional 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Additional programming by "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblcreated 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Initially created and developed by "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblversion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1.10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentPos As Long
Dim thanksHeight As Long

Dim thanks() As String
Dim thankscount As Long


Private Sub Form_Click()
10        Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
10        Unload Me
End Sub

Private Sub Form_Load()
10        lblversion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
          
20        BitBlt picThanks.hDC, 0, 0, picThanks.width, picThanks.height, Me.hDC, picThanks.Left, picThanks.Top, vbSrcCopy
30        picThanks.Refresh
          
40        Set picThanks.Font = txtThanks.Font
          
50        currentPos = picThanks.height
          
60        thanksHeight = picThanks.TextHeight(txtThanks.Text)
          
70        thanks = Split(txtThanks.Text, vbCrLf)
80        thankscount = UBound(thanks) + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        inSplash = False
End Sub



Private Sub lbladditional_Click()
10        Unload Me
End Sub

Private Sub lblcreated_Click()
10        Unload Me
End Sub





Private Sub lblEmailDrake_Click()
    ShellExecute 0&, vbNullString, "mailto:DCME.Continuum@gmail.com?subject=To Drake: ", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lblEmailSamapico_Click()
    ShellExecute 0&, vbNullString, "mailto:DCME.Continuum@gmail.com?subject=To Samapico: ", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lbllink_Click()
10        ShellExecute Me.hWnd, "open", "http://www.ssforum.net/index.php?showforum=277", _
                       vbNullString, vbNullString, SW_SHOWNORMAL
20        Unload Me

End Sub

Private Sub lblversion_Click()
10        Unload Me
End Sub

Private Sub lblvisitssforum_Click()
10        Unload Me
End Sub

Private Sub picThanks_Click()
10        Unload Me
End Sub

Private Sub Timer1_Timer()
      '    Static lastUpdate As Long
      '    Dim curTime As Long
      '
      '    curTime = GetTickCount
      '    Label2.Caption = curTime - lastUpdate
      '    lastUpdate = curTime
          Dim i As Long
          
10        picThanks.Cls
          
20        BitBlt picThanks.hDC, 0, 0, picThanks.width, picThanks.height, Me.hDC, picThanks.Left, picThanks.Top, vbSrcCopy
              
30        currentPos = currentPos - 1
          
40        picThanks.CurrentX = 0
50        picThanks.CurrentY = currentPos
          
          Dim linewidth As Integer
          Dim toprint As String
          
60        For i = 0 To thankscount - 1
              
70            toprint = thanks(i)
              
80            picThanks.FontBold = (InStr(toprint, "<b>") <> 0)
90            toprint = replace(toprint, "<b>", "")
              
100           picThanks.FontItalic = (InStr(toprint, "<i>") <> 0)
110           toprint = replace(toprint, "<i>", "")
                  
120           linewidth = picThanks.TextWidth(toprint)
130           picThanks.CurrentX = picThanks.width / 2 - linewidth / 2
140           picThanks.Print toprint
150       Next
          
160       picThanks.Refresh
          
          'Loop if finished
170       If currentPos <= -thanksHeight Then
180           currentPos = picThanks.height
190       End If
End Sub

Private Sub TimerClose_Timer()
10        Unload Me
End Sub
