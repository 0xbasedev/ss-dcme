VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLog 
   Caption         =   "Debug Log"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Log As..."
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      Height          =   1215
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmLog.frx":0000
      Top             =   600
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    Call ClearDebugLog
    txtLog.Text = ""
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errh
    Dim f As Integer
    
    cd.DialogTitle = "Save log as..."

    'ask for overwrite
    cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

    cd.filename = "DCME_log_" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & ".log"
    cd.Filter = "*.log|*.log|*.txt|*.txt"
    cd.ShowSave
    
    DeleteFile cd.filename
    
    f = FreeFile
    'Ouvre le fichier
    Open cd.filename For Output As #f
    'Ecrit le texte dans le fichier
    Print #f, txtLog.Text
    'Ferme le fichier
    Close #f

    Exit Sub

errh:
    If Err = cdlCancel Then
        Exit Sub
    Else
        MessageBox Err & " " & Err.description, vbCritical
    End If

End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    txtLog.Text = GetDebugLog
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Form_Resize()
    If Me.height < (cmdSave.height + 64) * Screen.TwipsPerPixelY Then
        Me.height = (cmdSave.height + 64) * Screen.TwipsPerPixelY
    End If
    If Me.width < 64 * Screen.TwipsPerPixelX Then
        Me.width = 64 * Screen.TwipsPerPixelX
    End If

    txtLog.Left = 0
    txtLog.Top = 0
    txtLog.width = Me.ScaleWidth
    txtLog.height = Me.ScaleHeight - cmdSave.height
    cmdSave.Left = 0
    cmdSave.Top = txtLog.height
    cmdSave.width = Me.ScaleWidth \ 2
    cmdClear.Left = cmdSave.width
    cmdClear.Top = txtLog.height
    cmdClear.width = Me.ScaleWidth \ 2


End Sub
