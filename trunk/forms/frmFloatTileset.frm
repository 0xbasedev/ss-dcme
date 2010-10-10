VERSION 5.00
Begin VB.Form frmFloatTileset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tileset"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTileset 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   0
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.PictureBox picwalltileprev 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   2400
         Picture         =   "frmFloatTileset.frx":0000
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Shape shptilesel 
         BorderColor     =   &H0000FFFF&
         Height          =   240
         Index           =   2
         Left            =   240
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shptilesel 
         BorderColor     =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFloatTileset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim autoHide As Boolean

Private Sub Form_Activate()
    BitBlt pictileset.hDC, 0, 0, 304, 224, frmGeneral.cTileset.Pic_Tileset.hDC, 0, 0, vbSrcCopy

    Call UpdateSelections
    autoHide = True
End Sub

Private Sub Form_Deactivate()
    If autoHide Then Me.visible = False
End Sub


Private Sub Form_Load()
    'AddDebug "Float tileset loaded"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub MakeSelectionFit(ByRef shape As shape, ByRef refshape As shape)
    'Gives the same position and size to 'shape'
    With shape
        If .Top <> refshape.Top Then .Top = refshape.Top
        If .Left <> refshape.Left Then .Left = refshape.Left
        If .width <> refshape.width Then .width = refshape.width
        If .height <> refshape.height Then .height = refshape.height
    End With
End Sub

Private Sub UpdateSelections()
          'TODO: FIX THIS
'          Dim i As Integer
'10        For i = vbLeftButton To vbRightButton
'20            MakeSelectionFit shptilesel(i), frmGeneral.shptilesel(i)
'30        Next
'
'40        If ShapesOverlap(shptilesel(1), shptilesel(2)) Then
'50            shptilesel(2).BorderColor = vbWhite
'60        Else
'70            shptilesel(2).BorderColor = vbYellow
'80        End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    'AddDebug "Float tileset unloaded"
End Sub

Private Sub pictileset_DblClick()
'10        Call frmGeneral.pictileset_DblClick
End Sub

Private Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'10        Call frmGeneral.pictileset_MouseDown(Button, Shift, X, Y)
'20        Call UpdateSelections
End Sub

Private Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'10        Call frmGeneral.pictileset_MouseMove(button, Shift, X, Y)
'20        Call UpdateSelections
End Sub

Private Sub pictileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'10        autoHide = False
'20        Call frmGeneral.pictileset_MouseUp(Button, Shift, X, Y)
'30        Call UpdateSelections
'40        Me.show
End Sub
