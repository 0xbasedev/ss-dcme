VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ToolProperty 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   128
   ToolboxBitmap   =   "ToolProperty.ctx":0000
   Begin VB.Frame frm 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   1935
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "999"
         Top             =   128
         Width           =   285
      End
      Begin MSComctlLib.Slider sld 
         Height          =   210
         Left            =   720
         TabIndex        =   2
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   370
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   128
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   64
         Value           =   1
      End
   End
End
Attribute VB_Name = "ToolProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Lboundval As Integer
Dim Uboundval As Integer
Dim pval As Integer
Dim cap As String

'Dim largechange As Integer

Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCaret& Lib "user32" (ByVal hWnd As Long)

Event Change()



Private Sub sld_Change()
    txt.Text = sld.value
    pval = sld.value

    RaiseEvent Change

End Sub

Private Sub sld_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    largechange = sld.largechange
'    sld.largechange = 1
'    Call sld_Change
'    RaiseEvent change
End Sub
Private Sub sld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button Then Call sld_MouseDown(Button, Shift, x, y)
End Sub

Private Sub sld_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    sld.largechange = largechange
End Sub

Private Sub sld_Scroll()
    Call sld_Change
End Sub

Private Sub txt_Change()
    Call removeDisallowedCharacters(txt, CSng(Lboundval), CSng(Uboundval), True)
    sld.value = val(txt.Text)
    pval = sld.value

    ' RaiseEvent Change
End Sub

Private Sub txt_Click()
    Call toggleLockToolTextBox(txt)
End Sub
Private Sub txt_LostFocus()
    Call toggleLockToolTextBox(txt, False)
End Sub

Public Property Get Min() As Integer
    Min = Lboundval
End Property

Public Property Let Min(ByVal newMin As Integer)
    Lboundval = newMin
    sld.Min = Lboundval
    If value < Lboundval Then
        value = Lboundval
    End If
    PropertyChanged "Min"
End Property

Public Property Get Max() As Integer
    Max = Uboundval
End Property

Public Property Let Max(ByVal newMax As Integer)
    Uboundval = newMax
    sld.Max = Uboundval
    If value > Uboundval Then
        value = Uboundval
    End If

    PropertyChanged "Max"
End Property

Public Property Get value() As Integer
    value = pval
End Property

Public Property Let value(ByVal newValue As Integer)
    If newValue < Lboundval Then
        pval = Lboundval
    Else
        pval = newValue
    End If

    If newValue > Uboundval Then
        pval = Uboundval
    Else
        pval = newValue
    End If

    sld.value = pval

    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
    Caption = cap
End Property

Public Property Let Caption(ByVal newCaption As String)
    cap = newCaption
    frm.Caption = cap
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Initialize()
    On Error Resume Next
    frm.Caption = cap
    Lboundval = 1
    sld.Min = 1
    Uboundval = 128
    sld.Max = 128
    value = 1
    sld.value = 1
    txt.Text = value
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    frm.Left = 0
    'frm.Top = 0
    Dim dWidth As Long
    dWidth = frm.width - (ScaleWidth - 1)
    frm.width = ScaleWidth - 1
    'frm.Height = ScaleHeight - 1

    'frm.Move 0, , ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight + Screen.TwipsPerPixelY * 4
    sld.Left = frm.width * Screen.TwipsPerPixelX - sld.width - Screen.TwipsPerPixelX * 2
    'sld.Move frm.Width - sld.Width - Screen.TwipsPerPixelX
    txt.Left = sld.Left - txt.width + Screen.TwipsPerPixelX * 3

    height = 345
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    Caption = Extender.name

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "")
    Lboundval = PropBag.ReadProperty("Min", 1)
    Uboundval = PropBag.ReadProperty("Max", 128)
    value = PropBag.ReadProperty("Value", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", Caption
    PropBag.WriteProperty "Min", Min
    PropBag.WriteProperty "Max", Max
    PropBag.WriteProperty "Value", value
End Sub



Private Sub removeDisallowedCharacters(ByRef txtbox As TextBox, lowerBound As Single, upperBound As Single, Optional dec As Boolean = False)
    If Not IsNumeric(txtbox.Text) Or InStr(txtbox.Text, "e") > 0 Or (Not dec And InStr(txtbox.Text, ".") > 0) Then
        Dim oldselstart As Integer
        oldselstart = txtbox.selstart - 1    'char  typed so always one more
        If oldselstart < 0 Then oldselstart = 0

        'remove all characters aside from nrs
        Dim i As Integer
        Dim finalresult As String
        For i = 1 To Len(txtbox.Text)
            If Asc(Mid$(txtbox.Text, i, 1)) < Asc("0") Or _
               Asc(Mid$(txtbox.Text, i, 1)) > Asc("9") Then
                Dim result As String
                If i - 1 >= 1 Then result = Mid$(txtbox.Text, 1, i - 1)
                If i + 1 <= Len(txtbox.Text) Then result = result + Mid$(txtbox.Text, i + 1, Len(txtbox.Text) - (i))
                finalresult = result
            End If
        Next
        txtbox.Text = finalresult
        If oldselstart > Len(txtbox.Text) Then
            txtbox.selstart = Len(txtbox.Text)
        Else
            txtbox.selstart = oldselstart
        End If
    End If

    If val(txtbox.Text) < lowerBound Then
        txtbox.Text = lowerBound
    End If

    If val(txtbox.Text) > upperBound Then
        txtbox.Text = upperBound
    End If

End Sub

Sub toggleLockToolTextBox(txt As TextBox, Optional lck As Boolean = False)
    If txt.locked And Not lck Then
        txt.locked = False
        txt.BorderStyle = vbFixedSingle
        txt.BackColor = vbWhite
        txt.Alignment = vbCenter
        Call ShowCaret(txt.hWnd)
        txt.selstart = 0
        txt.sellength = Len(txt.Text)
    Else
        txt.locked = True
        txt.BorderStyle = 0
        txt.BackColor = &H8000000F
        txt.Alignment = vbRightJustify
        Call HideCaret(txt.hWnd)
    End If

End Sub


