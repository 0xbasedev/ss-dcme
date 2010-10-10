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
10        tmpcurrenttip = CurrentTip
          
20        If chkRandom.value = vbChecked Then
30            Do While CurrentTip = tmpcurrenttip
40                CurrentTip = Int((tips.count * Rnd) + 1)
                  ' Or, you could cycle through the Tips in order
50            Loop
60        Else
70            CurrentTip = CurrentTip + 1
80            If tips.count < CurrentTip Then
90                CurrentTip = 1
100           End If
110       End If
          
          ' Show it.
120       frmTip.DisplayCurrentTip

End Sub

Private Sub chkLoadTipsAtStartup_Click()
      ' save whether or not this form should be displayed at startup
10        Call SetSetting("ShowTips", chkLoadTipsAtStartup.value)
20        Call SaveSettings
End Sub

Private Sub chkRandom_Click()
10        If chkRandom.value = vbChecked Then
20            cmdPreviousTip.Enabled = False
30        Else
40            cmdPreviousTip.Enabled = True
50        End If
60        Call SetSetting("ShowTipsRandom", chkRandom.value)
70        Call SaveSettings
End Sub

Private Sub cmdNextTip_Click()
10        DoNextTip
End Sub

Private Sub cmdOK_Click()
10        Unload Me
End Sub

Private Sub cmdPreviousTip_Click()
10        CurrentTip = CurrentTip - 1
20        If 1 > CurrentTip Then
30            CurrentTip = tips.count
40        End If

          
          ' Show it.
50        frmTip.DisplayCurrentTip
End Sub

Public Sub DisplayCurrentTip()
10        If tips.count > 0 Then
20            lblTipText.Caption = tips.item(CurrentTip)
30        End If
End Sub

Private Sub Form_Load()
'10        forceLoad = False
          
          Dim ShowAtStartup As Integer

20        Set Me.Icon = frmGeneral.Icon

'          ' See if we should be shown at startup
30        ShowAtStartup = GetSetting("ShowTips", 1)
'40        If ShowAtStartup = 0 And Not forceLoad Then
'50            Unload Me
'60            Exit Sub
'70        End If

80        chkLoadTipsAtStartup = ShowAtStartup

90        chkRandom.value = GetSetting("ShowTipsRandom", vbChecked)
100       If chkRandom.value = vbChecked Then
110           cmdPreviousTip.Enabled = False
120       Else
130           cmdPreviousTip.Enabled = True
              'Seed Rnd

131           Randomize
140       End If
          


          Dim i As Integer
160       For i = 0 To lstTips.ListCount - 1
170           If lstTips.list(i) <> "" Then
180               tips.add lstTips.list(i)
190           End If
200       Next


210       DoNextTip
End Sub
