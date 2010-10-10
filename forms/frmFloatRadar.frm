VERSION 5.00
Begin VB.Form frmFloatRadar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radar"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicRadar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   0
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "frmFloatRadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public autoHide As Boolean

Private Sub Form_Activate()
    MakeTopMost (Me.hWnd)
    Call UpdateRadar

    ' autoHide = True
    Debug.Print "activated"
End Sub

Private Sub Form_Deactivate()
    If autoHide Then
        Me.visible = False
    End If
    Debug.Print "deactivated with autohide " & autoHide
End Sub


Private Sub Form_Load()
    'AddDebug "Float radar loaded"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub


Private Sub UpdateRadar()
    BitBlt picradar.hDC, 0, 0, picradar.width, picradar.height, frmGeneral.picradar.hDC, 0, 0, vbSrcCopy
    picradar.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'AddDebug "Float radar unloaded"
End Sub

Private Sub picradar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    autoHide = False
    Debug.Print "mousedown"
    Me.show
    Call frmGeneral.picradar_MouseDown(Button, Shift, X, Y)
    Call UpdateRadar
    '   autoHide = True
    'Me.Show
    'Me.setfocus
End Sub

Private Sub picradar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    autoHide = False
    Call frmGeneral.picradar_MouseMove(Button, Shift, X, Y)
    Call UpdateRadar
    autoHide = True
    'Me.Show
End Sub

Private Sub picradar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeTopMost (Me.hWnd)
End Sub
