VERSION 5.00
Begin VB.Form frmPython 
   Caption         =   "Python Code Editor"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCode 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmPython.frx":0000
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmPython"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
10        If txtCode.Text <> frmGeneral.GetActiveRegionPythonCode Then
20            If MessageBox("Cancel without saving?", vbYesNo + vbQuestion) = vbOK Then
30                Unload Me
40            End If
50        Else
60            Unload Me
70        End If
End Sub

Private Sub cmdSave_Click()
10        Call frmGeneral.ChangeActiveRegionPythonCode(txtCode.Text)
20        DoEvents
30        Unload Me
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon

20        txtCode.Text = frmGeneral.GetActiveRegionPythonCode
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        cmdCancel_Click
End Sub

Private Sub Form_Resize()
10        If Me.height < cmdSave.height + 64 * Screen.TwipsPerPixelY Then
20            Me.height = cmdSave.height + 64 * Screen.TwipsPerPixelY
30        End If
40        If Me.width < 64 * Screen.TwipsPerPixelX Then
50            Me.width = 64 * Screen.TwipsPerPixelX
60        End If

70        txtCode.Left = 0
80        txtCode.Top = 0
90        txtCode.width = Me.ScaleWidth
100       txtCode.height = Me.ScaleHeight - cmdSave.height
110       cmdSave.Left = 0
120       cmdSave.Top = txtCode.height
130       cmdSave.width = Me.ScaleWidth \ 2
140       cmdCancel.Left = cmdSave.width
150       cmdCancel.Top = txtCode.height
160       cmdCancel.width = Me.ScaleWidth \ 2


End Sub

