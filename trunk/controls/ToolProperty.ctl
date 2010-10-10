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
         Name            =   "Arial Narrow"
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
            Name            =   "Arial Narrow"
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

Event change()



Private Sub sld_Change()
10        txt.Text = sld.value
20        pval = sld.value

30        RaiseEvent change

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
10        Call sld_Change
End Sub

Private Sub txt_Change()
10        Call removeDisallowedCharacters(txt, CSng(Lboundval), CSng(Uboundval), True)
20        sld.value = val(txt.Text)
30        pval = sld.value

          ' RaiseEvent Change
End Sub

Private Sub txt_Click()
10        Call toggleLockToolTextBox(txt)
End Sub
Private Sub txt_LostFocus()
10        Call toggleLockToolTextBox(txt, False)
End Sub

Public Property Get Min() As Integer
10        Min = Lboundval
End Property

Public Property Let Min(ByVal newMin As Integer)
10        Lboundval = newMin
20        sld.Min = Lboundval
30        If value < Lboundval Then
40            value = Lboundval
50        End If
60        PropertyChanged "Min"
End Property

Public Property Get Max() As Integer
10        Max = Uboundval
End Property

Public Property Let Max(ByVal newMax As Integer)
10        Uboundval = newMax
20        sld.Max = Uboundval
30        If value > Uboundval Then
40            value = Uboundval
50        End If

60        PropertyChanged "Max"
End Property

Public Property Get value() As Integer
10        value = pval
End Property

Public Property Let value(ByVal newValue As Integer)
10        If newValue < Lboundval Then
20            pval = Lboundval
30        Else
40            pval = newValue
50        End If

60        If newValue > Uboundval Then
70            pval = Uboundval
80        Else
90            pval = newValue
100       End If

110       sld.value = pval

120       PropertyChanged "Value"
End Property

Public Property Get Caption() As String
10        Caption = cap
End Property

Public Property Let Caption(ByVal newCaption As String)
10        cap = newCaption
20        frm.Caption = cap
30        PropertyChanged "Caption"
End Property

Private Sub UserControl_Initialize()
10        On Error Resume Next
20        frm.Caption = cap
30        Lboundval = 1
40        sld.Min = 1
50        Uboundval = 128
60        sld.Max = 128
70        value = 1
80        sld.value = 1
90        txt.Text = value
End Sub

Private Sub UserControl_Resize()
10        On Error Resume Next
20        frm.Left = 0
          'frm.Top = 0
          Dim dWidth As Long
30        dWidth = frm.width - (ScaleWidth - 1)
40        frm.width = ScaleWidth - 1
          'frm.Height = ScaleHeight - 1

          'frm.Move 0, , ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight + Screen.TwipsPerPixelY * 4
50        sld.Left = frm.width * Screen.TwipsPerPixelX - sld.width - Screen.TwipsPerPixelX * 2
          'sld.Move frm.Width - sld.Width - Screen.TwipsPerPixelX
60        txt.Left = sld.Left - txt.width + Screen.TwipsPerPixelX * 3

70        height = 345
End Sub

Private Sub UserControl_InitProperties()
10        On Error Resume Next
20        Caption = Extender.name

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
10        Caption = PropBag.ReadProperty("Caption", "")
20        Lboundval = PropBag.ReadProperty("Min", 1)
30        Uboundval = PropBag.ReadProperty("Max", 128)
40        value = PropBag.ReadProperty("Value", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
10        PropBag.WriteProperty "Caption", Caption
20        PropBag.WriteProperty "Min", Min
30        PropBag.WriteProperty "Max", Max
40        PropBag.WriteProperty "Value", value
End Sub



Private Sub removeDisallowedCharacters(ByRef txtbox As TextBox, lowerBound As Single, upperBound As Single, Optional dec As Boolean = False)
10        If Not IsNumeric(txtbox.Text) Or InStr(txtbox.Text, "e") > 0 Or (Not dec And InStr(txtbox.Text, ".") > 0) Then
              Dim oldselstart As Integer
20            oldselstart = txtbox.selstart - 1    'char  typed so always one more
30            If oldselstart < 0 Then oldselstart = 0

              'remove all characters aside from nrs
              Dim i As Integer
              Dim finalresult As String
40            For i = 1 To Len(txtbox.Text)
50                If Asc(Mid$(txtbox.Text, i, 1)) < Asc("0") Or _
                     Asc(Mid$(txtbox.Text, i, 1)) > Asc("9") Then
                      Dim result As String
60                    If i - 1 >= 1 Then result = Mid$(txtbox.Text, 1, i - 1)
70                    If i + 1 <= Len(txtbox.Text) Then result = result + Mid$(txtbox.Text, i + 1, Len(txtbox.Text) - (i))
80                    finalresult = result
90                End If
100           Next
110           txtbox.Text = finalresult
120           If oldselstart > Len(txtbox.Text) Then
130               txtbox.selstart = Len(txtbox.Text)
140           Else
150               txtbox.selstart = oldselstart
160           End If
170       End If

180       If val(txtbox.Text) < lowerBound Then
190           txtbox.Text = lowerBound
200       End If

210       If val(txtbox.Text) > upperBound Then
220           txtbox.Text = upperBound
230       End If

End Sub

Sub toggleLockToolTextBox(txt As TextBox, Optional lck As Boolean = False)
10        If txt.locked And Not lck Then
20            txt.locked = False
30            txt.BorderStyle = vbFixedSingle
40            txt.BackColor = vbWhite
50            txt.Alignment = vbCenter
60            Call ShowCaret(txt.hWnd)
70            txt.selstart = 0
80            txt.sellength = Len(txt.Text)
90        Else
100           txt.locked = True
110           txt.BorderStyle = 0
120           txt.BackColor = &H8000000F
130           txt.Alignment = vbRightJustify
140           Call HideCaret(txt.hWnd)
150       End If

End Sub


