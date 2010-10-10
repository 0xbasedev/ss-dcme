VERSION 5.00
Begin VB.UserControl cProgressBar 
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ToolboxBitmap   =   "cProgressBar.ctx":0000
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1605
      TabIndex        =   0
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape shpOutline 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Shape shpFill 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      Height          =   720
      Left            =   15
      Top             =   15
      Width           =   1935
   End
End
Attribute VB_Name = "cProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum displaymode
    dPercentage = 0 'automatically displays value as value/max (%)
    dBytes          'displays value as value kB / max kB, making the necessary conversions (bytes -> kB -> MB)
    dValue          'displays Value & caption
    dHidden         'no text
End Enum

'&H0000C000&&H00FF8080&

'Default Property Values:
Const m_def_FillColor = &HFF8080
Const m_def_FontColor = vbBlack
Const m_def_FillColorEnd = &HC000&
Const m_def_FontColorEnd = vbBlack
Const m_def_Value = 0
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_DisplayMode = dPercentage
Const m_def_DisplayDecimals = 1

Const m_def_Caption = ""

'Const m_def_FillColorEnd = vbBlue
Const m_def_UseGradientFill = True
'Const m_def_ForeColor = 0
'Const m_def_FillColor = 0
'Const m_def_Value = 0
'Const m_def_Min = 0
'Const m_def_Max = 100
'Property Variables:
Dim m_FillColor As OLE_COLOR
Dim m_FontColor As OLE_COLOR
Dim m_FillColorEnd As OLE_COLOR
Dim m_FontColorEnd As OLE_COLOR
'Dim m_FontColorEnd As Long
Dim m_Value As Long
Dim m_Min As Long
Dim m_Max As Long
Dim m_DisplayMode As displaymode
Dim m_DisplayDecimals As Integer

Dim m_Caption As String

'Dim m_FillColorEnd As Long
Dim m_UseGradientFill As Boolean
'Dim m_ForeColor As Long
'Dim m_Font As Font
'Dim m_FillColor As Long
'Dim m_Value As Variant
'Dim m_Min As Variant
'Dim m_Max As Variant
'Event Declarations:
Event change() 'MappingInfo=lblValue,lblValue,-1,Change
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."



Private Function BytesToFormat(bytes As Long) As String
      'Returns Max and formats it to Bytes, kB or MB when needed
10        If bytes < 1024 Then
20            BytesToFormat = bytes & " bytes"
30        ElseIf bytes < 1048576 Then '1024*1024
40            BytesToFormat = Format$(bytes / 1024, "0") & " KB"
50        Else
60            BytesToFormat = Format$(bytes / 1048576, "0.0") & " MB"
70        End If
          
End Function

'Event Change()
'
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get ForeColor() As Long
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As Long)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
10        Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
10        UserControl.Enabled() = New_Enabled
20        PropertyChanged "Enabled"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=6,0,0,0
'Public Property Get Font() As Font
'    Set Font = m_Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set m_Font = New_Font
'    PropertyChanged "Font"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
10        BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
10        UserControl.BackStyle() = New_BackStyle
20        PropertyChanged "BackStyle"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = UserControl.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    UserControl.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Private Sub ResetLabel()
10        Select Case displaymode
          
          Case dPercentage
20            lblValue.Caption = " %"
30        Case dBytes
40            lblValue.Caption = " / " & BytesToFormat(Max)
50        Case dValue
60            lblValue.Caption = " " & Caption
70        Case dHidden
80            lblValue.Caption = ""
90        End Select
          
100       lblValue.visible = (displaymode <> dHidden)
110       lblValue.Left = Me.ScaleWidth \ 2
120       lblValue.Top = Me.ScaleHeight \ 2 - lblValue.height \ 2
End Sub


Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
10        If m_Max <> 0 Then
              Dim newpercent As Double
20            newpercent = CDbl(value) / Max
              
              Dim percentvalue As Double
30            percentvalue = CDbl(value) * 100 / Max
              
40            If displaymode = dPercentage Then
50                If DisplayDecimals = 0 Then
                      '98 %
60                    lblValue.Caption = Int(percentvalue) & " %"
70                Else
                      '98.20 %
80                    lblValue.Caption = Format$(percentvalue, "0." & String$(DisplayDecimals, "0")) & " %"
90                End If
100           ElseIf displaymode = dBytes Then
110               lblValue.Caption = BytesToFormat(value) & " / " & BytesToFormat(Max)
                  
120           Else
130               lblValue.Caption = m_Value & " " & Caption
140           End If
              
              
              Dim newWidth As Single
              
150           newWidth = newpercent * shpOutline.width
              
160           If newWidth = 0 Then
170               shpFill.width = 1
180               shpFill.visible = False
190           Else
200               If shpFill.width <> newWidth Then shpFill.width = newWidth
                  
                  'Set correct color if gradient fill is used
210               If UseGradientFill Then
                      Dim newFillColor As Long
                      Dim newFontColor As Long
                          
220                   If Me.FillColor <> Me.FillColorEnd Then
230                       newFillColor = GetGradient(Me.FillColor, Me.FillColorEnd, newpercent)
240                       newFontColor = GetGradient(Me.FontColor, Me.FontColorEnd, newpercent)
250                   Else
260                       newFillColor = Me.FillColor
270                       newFontColor = Me.FontColor
280                   End If
                      
290                   shpFill.BackColor = newFillColor
300                   lblValue.ForeColor = newFontColor

310               Else
                      'Reset to normal colors
320                   If shpFill.FillColor <> Me.FillColor Then shpFill.FillColor = Me.FillColor
330                   If lblValue.ForeColor <> Me.FontColor Then lblValue.ForeColor = Me.FontColor
                      
340               End If
                  
350               shpFill.visible = True
360           End If
370       End If
End Sub


Private Function GetGradient(startColor As Long, endColor As Long, percent As Double) As Long
          'Returns a color between startColor and endColor
10        If percent > 1# Then percent = 1#
20        If percent < 0# Then percent = 0#
          
          Dim newR As Integer, newG As Integer, newB As Integer
          
30        newR = GetRED(startColor) + percent * (GetRED(endColor) - GetRED(startColor))
40        newG = GetGREEN(startColor) + percent * (GetGREEN(endColor) - GetGREEN(startColor))
50        newB = GetBLUE(startColor) + percent * (GetBLUE(endColor) - GetBLUE(startColor))
          
60        GetGradient = RGB(newR, newG, newB)
End Function



Private Sub UserControl_Click()
10        RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
10        RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get FillColor() As Long
'    FillColor = m_FillColor
'End Property
'
'Public Property Let FillColor(ByVal New_FillColor As Long)
'    m_FillColor = New_FillColor
'    PropertyChanged "FillColor"
'End Property
'


''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get Value() As Variant
'    Value = m_Value
'End Property
'
'Public Property Let Value(ByVal New_Value As Variant)
'    m_Value = New_Value
'
'    Refresh
'
'    PropertyChanged "Value"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get Min() As Variant
'    Min = m_Min
'End Property
'
'Public Property Let Min(ByVal New_Min As Variant)
'    m_Min = New_Min
'
'    Refresh
'
'    PropertyChanged "Min"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get Max() As Variant
'    Max = m_Max
'End Property
'
'Public Property Let Max(ByVal New_Max As Variant)
'    m_Max = New_Max
'
'    Refresh
'
'    PropertyChanged "Max"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
      '    m_ForeColor = m_def_ForeColor
      '    Set m_Font = Ambient.Font
      '    m_FillColor = m_def_FillColor
      '    m_Value = m_def_Value
      '    m_Min = m_def_Min
      '    m_Max = m_def_Max
10        m_Value = m_def_Value
20        m_Min = m_def_Min
30        m_Max = m_def_Max
      '    m_FillColorEnd = m_def_FillColorEnd
40        m_UseGradientFill = m_def_UseGradientFill
      '    m_FontColorEnd = m_def_FontColorEnd
          
50        shpFill.Left = 0
60        shpFill.Top = 0
          
          
70        m_FillColorEnd = m_def_FillColorEnd
80        m_FontColorEnd = m_def_FontColorEnd
          

90        m_FillColor = m_def_FillColor
100       m_FontColor = m_def_FontColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

      '    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
      '    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
10        UserControl.Enabled = PropBag.ReadProperty("Enabled", Vrai)
      '    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
20        UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)

30        UserControl.BackColor = PropBag.ReadProperty("BackColor", &H0&)
          
      '    lblValue.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
          
40        Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
50        Set lblValue.Font = PropBag.ReadProperty("Font", Ambient.Font)
          
      '    shpOutline.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
60        UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 4275)
70        UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
80        UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
90        UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 5910)
100       m_Value = PropBag.ReadProperty("Value", m_def_Value)
110       m_Min = PropBag.ReadProperty("Min", m_def_Min)
120       m_Max = PropBag.ReadProperty("Max", m_def_Max)
130       m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
          
140       m_UseGradientFill = PropBag.ReadProperty("UseGradientFill", m_def_UseGradientFill)
      '    lblValue.ForeColor = PropBag.ReadProperty("FontColor", &H80000012)
150       m_FillColorEnd = PropBag.ReadProperty("FillColorEnd", m_def_FillColorEnd)
160       m_FontColorEnd = PropBag.ReadProperty("FontColorEnd", m_def_FontColorEnd)
          

170       m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
180       m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
          
190       m_DisplayMode = PropBag.ReadProperty("DisplayMode", m_def_DisplayMode)
200       m_DisplayDecimals = PropBag.ReadProperty("DisplayDecimals", m_def_DisplayDecimals)
          
          
210       UserControl_Resize
End Sub

Private Sub UserControl_Resize()
10        shpOutline.width = Me.ScaleWidth
20        shpOutline.height = Me.ScaleHeight
          
30        shpFill.height = shpOutline.height
          
40        Call ResetLabel
          
50        Refresh
          
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

      '    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
      '    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
10        Call PropBag.WriteProperty("Enabled", UserControl.Enabled, Vrai)
      '    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
20        Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
      '    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
      '    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
      '    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
      '    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
30        Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H0&)
      '    Call PropBag.WriteProperty("ForeColor", lblValue.ForeColor, &H80000012)
      '    Call PropBag.WriteProperty("FillColor", shpFill.FillColor, &H0&)
40        Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
50        Call PropBag.WriteProperty("Font", lblValue.Font, Ambient.Font)
      '    Call PropBag.WriteProperty("BorderWidth", shpOutline.BorderWidth, 1)
60        Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 4275)
70        Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
80        Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
90        Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 5910)
100       Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
110       Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
120       Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
130       Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
          
      '    Call PropBag.WriteProperty("FillColorEnd", m_FillColorEnd, m_def_FillColorEnd)
140       Call PropBag.WriteProperty("UseGradientFill", m_UseGradientFill, m_def_UseGradientFill)
      '    Call PropBag.WriteProperty("FontColorEnd", m_FontColorEnd, m_def_FontColorEnd)
      '    Call PropBag.WriteProperty("FontColor", lblValue.ForeColor, &H80000012)
150       Call PropBag.WriteProperty("FillColorEnd", m_FillColorEnd, m_def_FillColorEnd)
160       Call PropBag.WriteProperty("FontColorEnd", m_FontColorEnd, m_def_FontColorEnd)
170       Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
180       Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
          
190       Call PropBag.WriteProperty("DisplayMode", m_DisplayMode, m_def_DisplayMode)
200       Call PropBag.WriteProperty("DisplayDecimals", m_DisplayDecimals, m_def_DisplayDecimals)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpOutline,shpOutline,-1,FillColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
10        BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
10        UserControl.BackColor() = New_BackColor
20        PropertyChanged "BackColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=lblValue,lblValue,-1,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = lblValue.ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    lblValue.ForeColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property

'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=shpFill,shpFill,-1,FillColor
'Public Property Get FillColor() As OLE_COLOR
'    FillColor = shpFill.FillColor
'End Property
'
'Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
'    shpFill.FillColor() = New_FillColor
'    PropertyChanged "FillColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblValue,lblValue,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
10        Set Font = lblValue.Font
End Property

Public Property Set Font(ByVal New_Font As Font)

10        Set lblValue.Font = New_Font
              
20        Call ResetLabel
30        Refresh
          
40        PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=shpOutline,shpOutline,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = shpOutline.BorderStyle
'End Property

Private Sub lblValue_Change()
10        RaiseEvent change
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=shpOutline,shpOutline,-1,BorderWidth
'Public Property Get BorderWidth() As Integer
'    BorderWidth = shpOutline.BorderWidth
'End Property
'
'Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
'    shpOutline.BorderWidth() = New_BorderWidth
'    PropertyChanged "BorderWidth"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
10        ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
10        UserControl.ScaleHeight() = New_ScaleHeight
20        PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
10        ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
10        UserControl.ScaleLeft() = New_ScaleLeft
20        PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
10        ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
10        UserControl.ScaleTop() = New_ScaleTop
20        PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
10        ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
10        UserControl.ScaleWidth() = New_ScaleWidth
20        PropertyChanged "ScaleWidth"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
10        Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
10        m_Caption = New_Caption
          
20        PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(ByVal str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
10        TextWidth = UserControl.TextWidth(str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(ByVal str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
10        TextHeight = UserControl.TextHeight(str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get value() As Long
10        value = m_Value
End Property

Public Property Let value(ByVal New_Value As Long)
10        m_Value = New_Value
          
20        Refresh
          
30        PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
10        Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
10        m_Min = New_Min
          
20        Refresh
          
30        PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
10        Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
10        If New_Max = 0 Then New_Max = 1
          
20        m_Max = New_Max
          
30        Call ResetLabel
40        Refresh
          
50        PropertyChanged "Max"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get FillColorEnd() As Long
'    FillColorEnd = m_FillColorEnd
'End Property
'
'Public Property Let FillColorEnd(ByVal New_FillColorEnd As Long)
'    m_FillColorEnd = New_FillColorEnd
'
'    If UseGradientFill Then Refresh
'
'    PropertyChanged "FillColorEnd"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseGradientFill() As Boolean
10        UseGradientFill = m_UseGradientFill
End Property

Public Property Let UseGradientFill(ByVal New_UseGradientFill As Boolean)
10        m_UseGradientFill = New_UseGradientFill
          
20        Refresh
          
30        PropertyChanged "UseGradientFill"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get FontColorEnd() As Long
'    FontColorEnd = m_FontColorEnd
'End Property
'
'Public Property Let FontColorEnd(ByVal New_FontColorEnd As Long)
'    m_FontColorEnd = New_FontColorEnd
'
'    If UseGradientFill Then Refresh
'
'    PropertyChanged "FontColorEnd"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=lblValue,lblValue,-1,ForeColor
'Public Property Get FontColor() As OLE_COLOR
'    FontColor = lblValue.ForeColor
'End Property
'
'Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
'    lblValue.ForeColor() = New_FontColor
'    PropertyChanged "FontColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0&
Public Property Get FillColorEnd() As OLE_COLOR
Attribute FillColorEnd.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
10        FillColorEnd = m_FillColorEnd
End Property

Public Property Let FillColorEnd(ByVal New_FillColorEnd As OLE_COLOR)
10        m_FillColorEnd = New_FillColorEnd
          
20        If UseGradientFill Then Refresh
          
30        PropertyChanged "FillColorEnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FontColorEnd() As OLE_COLOR
10        FontColorEnd = m_FontColorEnd
End Property

Public Property Let FontColorEnd(ByVal New_FontColorEnd As OLE_COLOR)
10        m_FontColorEnd = New_FontColorEnd
          
20        If UseGradientFill Then Refresh
          
30        PropertyChanged "FontColorEnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
10        FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
10        m_FillColor = New_FillColor
          
20        Refresh
          
30        PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,
Public Property Get FontColor() As OLE_COLOR
Attribute FontColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
10        FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
10        m_FontColor = New_FontColor
          
20        Refresh
          
30        PropertyChanged "FontColor"
End Property


Public Property Get displaymode() As displaymode
10        displaymode = m_DisplayMode
End Property

Public Property Let displaymode(ByVal New_DisplayMode As displaymode)
10        m_DisplayMode = New_DisplayMode
          
20        Call ResetLabel
30        Refresh
          
40        PropertyChanged "DisplayMode"
End Property

Public Property Get DisplayDecimals() As Integer
10        DisplayDecimals = m_DisplayDecimals
End Property

Public Property Let DisplayDecimals(ByVal New_DisplayDecimals As Integer)
10        m_DisplayDecimals = New_DisplayDecimals
          
20        Refresh
          
30        PropertyChanged "DisplayDecimals"
End Property

Public Property Get Text() As String
10        Text = lblValue.Caption
End Property
