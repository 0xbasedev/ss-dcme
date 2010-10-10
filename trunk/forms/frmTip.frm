VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   3285
   ClientLeft      =   2475
   ClientTop       =   2475
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdPreviousTip 
      Caption         =   "&Previous Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkRandom 
      Caption         =   "Random Order"
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   2940
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ListBox lstTips 
      Height          =   1815
      ItemData        =   "frmTip.frx":0000
      Left            =   480
      List            =   "frmTip.frx":0073
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2940
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":1215
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   1
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim tips As New Collection

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()
    Dim tmpcurrenttip As Long
    ' Select a tip at random.
    tmpcurrenttip = CurrentTip
    
    If chkRandom.value = vbChecked Then
        Do While CurrentTip = tmpcurrenttip
            CurrentTip = Int((tips.count * Rnd) + 1)
            ' Or, you could cycle through the Tips in order
        Loop
    Else
        CurrentTip = CurrentTip + 1
        If tips.count < CurrentTip Then
            CurrentTip = 1
        End If
    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip

End Sub

Private Sub chkLoadTipsAtStartup_Click()
' save whether or not this form should be displayed at startup
    Call SetSetting("ShowTips", chkLoadTipsAtStartup.value)
    Call SaveSettings
End Sub

Private Sub chkRandom_Click()
    If chkRandom.value = vbChecked Then
        cmdPreviousTip.Enabled = False
    Else
        cmdPreviousTip.Enabled = True
    End If
    Call SetSetting("ShowTipsRandom", chkRandom.value)
    Call SaveSettings
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPreviousTip_Click()
    CurrentTip = CurrentTip - 1
    If 1 > CurrentTip Then
        CurrentTip = tips.count
    End If

    
    ' Show it.
    frmTip.DisplayCurrentTip
End Sub

Public Sub DisplayCurrentTip()
    If tips.count > 0 Then
        lblTipText.Caption = tips.item(CurrentTip)
    End If
End Sub

Private Sub Form_Load()
'10        forceLoad = False
    
    Dim ShowAtStartup As Integer

    Set Me.Icon = frmGeneral.Icon

'          ' See if we should be shown at startup
    ShowAtStartup = GetSetting("ShowTips", 1)
'40        If ShowAtStartup = 0 And Not forceLoad Then
'50            Unload Me
'60            Exit Sub
'70        End If

    chkLoadTipsAtStartup = ShowAtStartup

    chkRandom.value = GetSetting("ShowTipsRandom", vbChecked)
    If chkRandom.value = vbChecked Then
        cmdPreviousTip.Enabled = False
    Else
        cmdPreviousTip.Enabled = True
        'Seed Rnd

        Randomize
    End If
    


    Dim i As Integer
    For i = 0 To lstTips.ListCount - 1
        If lstTips.list(i) <> "" Then
            tips.add lstTips.list(i)
        End If
    Next


    DoNextTip
End Sub
