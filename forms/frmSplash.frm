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
      Top             =   120
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
      Left            =   3840
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbllink 
      Alignment       =   2  'Center
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
      Caption         =   "Visit                          for info and updates"
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
   Begin VB.Label lbland 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblcreated 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Created by"
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
      Left            =   1320
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   120
      Width           =   2295
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
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblversion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    BitBlt picThanks.hDC, 0, 0, picThanks.width, picThanks.height, Me.hDC, picThanks.Left, picThanks.Top, vbSrcCopy
    picThanks.Refresh
    
    Set picThanks.Font = txtThanks.Font
    
    currentPos = picThanks.height
    
    thanksHeight = picThanks.TextHeight(txtThanks.Text)
    
    thanks = Split(txtThanks.Text, vbCrLf)
    thankscount = UBound(thanks) + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    inSplash = False
End Sub





Private Sub lblcreated_Click()
    Unload Me
End Sub





Private Sub lblEmailDrake_Click()
    ShellExecute 0&, vbNullString, "mailto:DCME.Continuum@gmail.com?subject=To Drake: ", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lblEmailSamapico_Click()
    ShellExecute 0&, vbNullString, "mailto:DCME.Continuum@gmail.com?subject=To Samapico: ", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lbllink_Click()
    ShellExecute Me.hWnd, "open", "http://www.ssforum.net/index.php?showforum=277", _
                 vbNullString, vbNullString, SW_SHOWNORMAL
    Unload Me

End Sub

Private Sub lblversion_Click()
    Unload Me
End Sub

Private Sub lblvisitssforum_Click()
    Unload Me
End Sub

Private Sub picThanks_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
'    Static lastUpdate As Long
'    Dim curTime As Long
'
'    curTime = GetTickCount
'    Label2.Caption = curTime - lastUpdate
'    lastUpdate = curTime
    Dim i As Long
    
    picThanks.Cls
    
    BitBlt picThanks.hDC, 0, 0, picThanks.width, picThanks.height, Me.hDC, picThanks.Left, picThanks.Top, vbSrcCopy
        
    currentPos = currentPos - 1
    
    picThanks.CurrentX = 0
    picThanks.CurrentY = currentPos
    
    Dim linewidth As Integer
    Dim toprint As String
    
    For i = 0 To thankscount - 1
        
        toprint = thanks(i)
        
        picThanks.FontBold = (InStr(toprint, "<b>") <> 0)
        toprint = replace(toprint, "<b>", "")
        
        picThanks.FontItalic = (InStr(toprint, "<i>") <> 0)
        toprint = replace(toprint, "<i>", "")
            
        linewidth = picThanks.TextWidth(toprint)
        picThanks.CurrentX = picThanks.width / 2 - linewidth / 2
        picThanks.Print toprint
    Next
    
    picThanks.Refresh
    
    'Loop if finished
    If currentPos <= -thanksHeight Then
        currentPos = picThanks.height
    End If
End Sub

Private Sub TimerClose_Timer()
    Unload Me
End Sub
