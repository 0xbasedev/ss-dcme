VERSION 5.00
Begin VB.Form dlgCancelWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Waiting..."
   ClientHeight    =   915
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540.612
   ScaleMode       =   0  'User
   ScaleWidth      =   3281.616
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "Waiting for editing program to end..."
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3360
   End
End
Attribute VB_Name = "dlgCancelWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents modifcheck As clsCheckModifs
Attribute modifcheck.VB_VarHelpID = -1

Dim myFile As String

Private Sub cmdCancel_Click()
    If ConfirmUnload Then
        CancelWait = True
'        Unload Me
    End If
End Sub

Private Function ConfirmUnload() As Boolean
    ConfirmUnload = (MessageBox("Any changes made in the remote application will be lost, are you sure you want to cancel?", vbYesNo + vbExclamation) = vbYes)
End Function

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon
    
    Set modifcheck = New clsCheckModifs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> QueryUnloadConstants.vbFormCode Then Cancel = Not ConfirmUnload
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set modifcheck = Nothing
    CancelWait = True
End Sub



Private Sub modifcheck_FileAction(filename As String, FA As FILE_ACTION)
    MsgBox "Action " & FA & " on file '" & filename & "'"
End Sub

Private Sub modifcheck_FileDeleted(filename As String)
    MsgBox "File '" & filename & "' deleted"
End Sub

Private Sub modifcheck_FileModified(filename As String)
    MsgBox "File '" & filename & "' modified"
End Sub
