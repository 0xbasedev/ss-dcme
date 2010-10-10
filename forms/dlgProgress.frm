VERSION 5.00
Begin VB.Form dlgProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress"
   ClientHeight    =   780
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   4575
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   StartUpPosition =   1  'CenterOwner
   Begin DCME.cProgressBar bar 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   33
      ScaleWidth      =   305
      DisplayDecimals =   0
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   4455
   End
End
Attribute VB_Name = "dlgProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Operation As String

Public Sub InitProgressBar(sOperation As String, Max As Long)
      'Initializes the progress bar
          
10        Operation = sOperation
20        bar.Max = Max
30        bar.value = 0
40        lblAction.Caption = ""
50        Call RefreshCaption
End Sub

Sub SetValue(value As Long)
      'Set the value of the progress bar
          
10        bar.value = value
End Sub

Sub SetOperation(sOperation As String)
      'Set the operation name
          
10        Operation = sOperation
20        Call RefreshCaption
End Sub

Sub SetLabel(newLabel As String)
      'Set the progressbar caption

10        lblAction.Caption = newLabel
End Sub

Private Sub RefreshCaption()
      'Refreshes the caption

10        Me.Caption = bar.Text & " - " & Operation & "..."
End Sub

Private Sub bar_change()
      'Update caption
10        Call RefreshCaption
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon
End Sub
