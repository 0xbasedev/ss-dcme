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
    If txtCode.Text <> frmGeneral.GetActiveRegionPythonCode Then
        If MessageBox("Cancel without saving?", vbYesNo + vbQuestion) = vbOK Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdSave_Click()
    Call frmGeneral.ChangeActiveRegionPythonCode(txtCode.Text)
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    txtCode.Text = frmGeneral.GetActiveRegionPythonCode
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click
End Sub

Private Sub Form_Resize()
    If Me.height < cmdSave.height + 64 * Screen.TwipsPerPixelY Then
        Me.height = cmdSave.height + 64 * Screen.TwipsPerPixelY
    End If
    If Me.width < 64 * Screen.TwipsPerPixelX Then
        Me.width = 64 * Screen.TwipsPerPixelX
    End If

    txtCode.Left = 0
    txtCode.Top = 0
    txtCode.width = Me.ScaleWidth
    txtCode.height = Me.ScaleHeight - cmdSave.height
    cmdSave.Left = 0
    cmdSave.Top = txtCode.height
    cmdSave.width = Me.ScaleWidth \ 2
    cmdCancel.Left = cmdSave.width
    cmdCancel.Top = txtCode.height
    cmdCancel.width = Me.ScaleWidth \ 2


End Sub

