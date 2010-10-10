VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PropertyList 
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ScaleHeight     =   3120
   ScaleWidth      =   8160
   ToolboxBitmap   =   "PropertyList.ctx":0000
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstChoice 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   5503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "PropertyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function setParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Const SW_SHOWNOACTIVATE = 4

Const ALLOWED_EXPRESSION_CHARACTERS = "01234567890-+()*/\%^."

Enum propertyType
    p_text
    p_number
    p_list
    p_expression 'expression that can be evaluated
End Enum

Private Type propstruct
    name As String
    value As String
    Type As propertyType
    choices() As String
    lbnd As Long
    ubnd As Long
    locked As Boolean
    tooltip As String
End Type

Dim props() As propstruct
Dim propcount As Integer

Dim curX As Integer
Dim curY As Integer
Dim curlstItem As ListItem

Dim edit As Boolean

Event PropertyChanged(propName As String)



Private Sub lst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim overlstItem As ListItem
          
10        Set overlstItem = lst.HitTest(X, Y)
          
20        If overlstItem Is Nothing Then
30            lst.tooltiptext = ""
40        Else
50            lst.tooltiptext = props(overlstItem.Index - 1).tooltip
60        End If
End Sub

Private Sub lst_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If edit Then
20            EditOff (True)
30            Exit Sub
40        End If
          
50        curX = X
60        curY = Y
70        Call EditOn
End Sub

Private Sub lstChoice_Click()
10        If edit Then EditOff (True)
          
End Sub

Private Sub txt_Change()
10        If edit And props(curlstItem.Index - 1).Type = p_number Then
20            Call removeDisallowedCharacters(txt, CSng(props(curlstItem.Index - 1).lbnd), CSng(props(curlstItem.Index - 1).ubnd), False)
30        End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
10        If KeyAscii = vbKeyReturn Then EditOff (True)
20        If KeyAscii = vbKeyEscape Then EditOff (False)
              
End Sub

Private Sub UserControl_Initialize()
10        ReDim props(0)
20        propcount = 0
End Sub

Private Sub UserControl_LostFocus()
10        If edit Then EditOff (True)
End Sub

Private Sub UserControl_Resize()
10        lst.width = UserControl.width
20        lst.height = UserControl.height
          
30        If propcount > 0 Then
40            If (propcount + 1) * lst.height > UserControl.height Then
              'scrollbar will appear
50                lst.ColumnHeaders(1).width = (lst.width / 2) - 9 * Screen.TwipsPerPixelX
60                lst.ColumnHeaders(2).width = lst.width / 2 - 9 * Screen.TwipsPerPixelX
70            Else
80                lst.ColumnHeaders(1).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
90                lst.ColumnHeaders(2).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
100           End If
110       Else
120           lst.ColumnHeaders(1).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
130           lst.ColumnHeaders(2).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
140       End If
End Sub

Sub AddProperty(name As String, proptype As propertyType, Optional lbnd As Long = -2147483647, Optional ubnd As Long = 2147483647)
10        If propcount > UBound(props) Then
20            ReDim Preserve props(UBound(props) + 5)
30        End If
          
          Dim val As Integer
40        val = getPropertyIndex(name)
50        If val <> -1 Then
60            Call Err.Raise(10002, , "Property '" & name & "' already exists")
70            Exit Sub
80        End If
          
90        props(propcount).name = name
100       props(propcount).Type = proptype
          
110       If proptype = p_number Then
120           props(propcount).value = 0
130           props(propcount).lbnd = lbnd
140           props(propcount).ubnd = ubnd
150       ElseIf proptype = p_list Then
160           ReDim props(propcount).choices(0)
170       End If
          
180       If lst.Enabled Then Call lst.ListItems.add(propcount + 1, name, name)
          
190       If (propcount + 1) * lst.ListItems(name).height > UserControl.height Then
              'scrollbar will appear
200           lst.ColumnHeaders(1).width = (lst.width / 2) - 9 * Screen.TwipsPerPixelX
210           lst.ColumnHeaders(2).width = lst.width / 2 - 9 * Screen.TwipsPerPixelX
220       End If
                  
          
230       propcount = propcount + 1
End Sub

Sub setPropertyToolTipText(name As String, tooltiptext As String)
          Dim idx As Integer
10        idx = getPropertyIndex(name)
          
20        If idx = -1 Then
30            Call Err.Raise(10000, , "Property '" & name & "' does not exist")
40            Exit Sub
50        End If
          
60        props(idx).tooltip = tooltiptext
End Sub

Sub setPropertyChoiceList(name As String, chlst() As String)
          Dim idx As Integer
10        idx = getPropertyIndex(name)
              
20        If idx = -1 Then
30            Call Err.Raise(10000, , "Property '" & name & "' does not exist")
40            Exit Sub
50        End If
          
60        If props(idx).Type = p_list Then
70            props(idx).choices = chlst
              
80            props(idx).value = 0
              'check if the current value is in the choice list, if it isn't, take the first choice
              'Dim i As Integer
              'Dim found As Boolean
              'For i = 0 To UBound(props(idx).choices)
              '    If props(idx).value = props(idx).choices(i) Then
              '        found = True
              '        Exit For
              '    End If
              'Next
              'If Not found Then
              '    props(idx).value = props(idx).choices(0)
              '
              '    If lst.Enabled Then lst.ListItems(idx + 1).SubItems(1) = props(idx).value
              'End If
90        Else
100           Call Err.Raise(10001, , "Property is not a list type")
110       End If
End Sub

Sub setPropertyNumberBoundaries(name As String, lbnd As Long, ubnd As Long)
          Dim idx As Integer
10        idx = getPropertyIndex(name)
          
20        If idx = -1 Then
30            Call Err.Raise(10000, , "Property '" & name & "' does not exist")
40            Exit Sub
50        End If
          
60        If props(idx).Type = p_number Then
70            props(idx).lbnd = lbnd
80            props(idx).ubnd = ubnd
90        Else
100           Call Err.Raise(10001, "Property is not a number type")
110       End If
End Sub

Sub setPropertyValue(name As String, value As String)
          
          Dim idx As Integer
10        idx = getPropertyIndex(name)
          
20        If idx = -1 Then
30            Call Err.Raise(10000, , "Property '" & name & "' does not exist")
40            Exit Sub
50        End If
          
60        If props(idx).Type = p_list Then
70            props(idx).value = val(value)
80        End If
          
90        If lst.Enabled Then
100           If props(idx).Type = p_list Then
110               lst.ListItems(props(idx).name).SubItems(1) = props(idx).choices(value)
120           Else
130               lst.ListItems(props(idx).name).SubItems(1) = value
140           End If
150       End If
End Sub

Sub setPropertyLocked(name As String, locked As Boolean)
          Dim idx As Integer
10        idx = getPropertyIndex(name)
          
20        If idx = -1 Then
30            Call Err.Raise(10000, , "Property '" & name & "' does not exist")
40            Exit Sub
50        End If
          
60        props(idx).locked = locked
End Sub

Function getPropertyValue(name As String) As String
          Dim idx As Integer
10        idx = getPropertyIndex(name)
          
20        If idx = -1 Then
30            Call Err.Raise(10000, , "Property '" & name & "' does not exist")
40            Exit Function
50        End If
          
60        getPropertyValue = props(idx).value
End Function

Sub RemoveProperty(name As String)
10        If propcount < UBound(props) - 5 Then
20            ReDim Preserve props(UBound(props) - 5)
30        End If
          
          Dim idx As Integer
40        idx = getPropertyIndex(name)
          
50        If idx <> -1 Then
              Dim i As Integer
60            For i = idx + 1 To propcount - 1
70                props(i - 1) = props(i)
80            Next
90        End If

End Sub

Private Function getPropertyIndex(name As String) As Integer
          Dim i As Integer
10        For i = 0 To propcount - 1
20            If props(i).name = name Then
30                getPropertyIndex = i
40                Exit Function
50            End If
60        Next
          
70        getPropertyIndex = -1
End Function

Sub EditOn()
10        Set curlstItem = lst.HitTest(curX, curY)
          
20        If curlstItem Is Nothing Then
30            Exit Sub
40        End If
          
50        If props(curlstItem.Index - 1).locked Then
60            Exit Sub
70        End If
          
80        Select Case props(curlstItem.Index - 1).Type
          
              Case p_text, p_number
90                txt.width = lst.ColumnHeaders(2).width - 2 * Screen.TwipsPerPixelX
100               txt.height = curlstItem.height - 1 * Screen.TwipsPerPixelY
110               txt.Left = lst.Left + curlstItem.Left + lst.ColumnHeaders(1).width
120               txt.Top = lst.Top + curlstItem.Top + 2 * Screen.TwipsPerPixelY
130               txt.Text = curlstItem.SubItems(1)
140               txt.selstart = 0
150               txt.sellength = Len(txt.Text)
160               txt.visible = True
170               txt.setfocus
                  
180           Case p_list

190               lstChoice.Left = lst.Left + curlstItem.Left + lst.ColumnHeaders(1).width
200               lstChoice.Top = lst.Top + curlstItem.Top + 1 * Screen.TwipsPerPixelY + curlstItem.height
210               lstChoice.width = lst.ColumnHeaders(2).width - 1 * Screen.TwipsPerPixelX

                  
220               lstChoice.Clear
                  Dim i As Integer
                  
230               For i = 0 To UBound(props(curlstItem.Index - 1).choices)
240                   If props(curlstItem.Index - 1).choices(i) <> "" Then
250                       lstChoice.addItem props(curlstItem.Index - 1).choices(i)
260                   End If
                  '    If lstChoice.list(lstChoice.Listcount - 1) = props(curlstItem.Index - 1).value Then
                 '         lstChoice.ListIndex = i
                 '     End If
270               Next
280               lstChoice.ListIndex = val(props(curlstItem.Index - 1).value)
                  
290               If lstChoice.ListIndex = -1 Then lstChoice.ListIndex = 0

300               If lstChoice.ListCount < 5 Then
310                   lstChoice.height = (curlstItem.height + 1) * lstChoice.ListCount
320               Else
330                   lstChoice.height = curlstItem.height * 5
340               End If
                  
                  
                  'show the lstchoice, give it to the desktopwindow, and move it to the right position
                  Dim pt As POINTAPI
350               pt.X = ScaleX(lstChoice.Left, vbTwips, vbPixels) '+ UserControl.parent.Left
360               pt.Y = ScaleY(lstChoice.Top, vbTwips, vbPixels)
370               ClientToScreen UserControl.hWnd, pt
                  'pt.x = pt.x + lstChoice.Left '+ UserControl.parent.Left
                  'pt.y = pt.y + lstChoice.Top
380               setParent lstChoice.hWnd, GetDesktopWindow
390               lstChoice.Move ScaleX(pt.X, vbPixels, vbTwips), ScaleY(pt.Y, vbPixels, vbTwips)
                  'lstChoice.Left = pt.x
                  'lstChoice.Top = pt.y
                  'MakeTopMost lstChoice.hWnd
                  'RestoreWin lstChoice.hWnd, False
                  'ShowWindow lstChoice.hWnd, SW_SHOWNORMAL
400               ShowWindow lstChoice.hWnd, SW_SHOWNOACTIVATE
410               BringWindowToTop lstChoice.hWnd
420               lstChoice.visible = True

                  'ShowWindow , SW_SHOWNORMAL
                  'lstChoice.setfocus
                  
                  
430       End Select
          
440       edit = True
End Sub

Sub EditOff(apply As Boolean)
          Dim changed As Boolean
          
10        If Not edit Then Exit Sub
          
20        Select Case props(curlstItem.Index - 1).Type
              Case p_text, p_number
                              
30                If apply Then
40                    If props(curlstItem.Index - 1).value <> txt.Text Then changed = True
50                    props(curlstItem.Index - 1).value = txt.Text
60                End If
                      
70                curlstItem.SubItems(1) = txt.Text
80                txt.visible = False
90                If changed Then RaiseEvent PropertyChanged(props(curlstItem.Index - 1).name)
                  
100           Case p_list
110               If apply Then
120                   If props(curlstItem.Index - 1).value <> lstChoice.ListIndex Then changed = True
130                   props(curlstItem.Index - 1).value = lstChoice.ListIndex
140               End If
150               curlstItem.SubItems(1) = lstChoice.list(lstChoice.ListIndex)
160               lstChoice.visible = False
                  'show the lstchoice, give it to the desktopwindow, and move it to the right position
170               setParent lstChoice.hWnd, UserControl.hWnd
180               lstChoice.Left = 0
190               lstChoice.Top = 0
                  'ShowWindow lstChoice.hWnd, SW_SHOWNOACTIVATE
                  
                  
200               If changed Then RaiseEvent PropertyChanged(props(curlstItem.Index - 1).name)
                  
210       End Select
220       edit = False
          
End Sub

Sub UpdateList()
10        lst.ListItems.Clear
          
          Dim i As Integer
20        For i = 0 To propcount - 1
30            Call lst.ListItems.add(, props(i).name, props(i).name)
40            If props(i).Type = p_list Then
50                lst.ListItems.item(props(i).name).SubItems(1) = props(i).choices(props(i).value)
60            Else
70                lst.ListItems.item(props(i).name).SubItems(1) = props(i).value
80            End If
90        Next
          
End Sub

'It is in general.bas
'Private Sub removeDisallowedCharacters(ByRef txtbox As TextBox, lowerBound As Single, upperBound As Single, Optional dec As Boolean = False)
'    If (Not IsNumeric(txtbox.text) And (txtbox.text <> "-" Or lowerBound >= 0)) _
'        Or InStr(txtbox.text, "e") > 0 Or InStr(txtbox.text, "E") > 0 _
'        Or (Not dec And (InStr(txtbox.text, ".") > 0 Or InStr(txtbox.text, ",") > 0)) _
'        Or (lowerBound < 0 And InStr(2, txtbox.text, "-") > 1) _
'        Or (lowerBound >= 0 And InStr(txtbox.text, "-") > 0) Then
'
'        Dim oldselstart As Integer
'        oldselstart = txtbox.selstart - 1    'char  typed so always one more
'        If oldselstart < 0 Then oldselstart = 0
'
'        'remove all characters aside from nrs
'        Dim i As Integer
'        Dim finalresult As String
'        For i = 1 To Len(txtbox.text)
'            If (Asc(Mid$(txtbox.text, i, 1)) < Asc("0") Or _
'                Asc(Mid$(txtbox.text, i, 1)) > Asc("9")) Then
'                Dim result As String
'                If i - 1 >= 1 Then result = Mid$(txtbox.text, 1, i - 1)
'                If i + 1 <= Len(txtbox.text) Then result = result + Mid$(txtbox.text, i + 1, Len(txtbox.text) - (i))
'                finalresult = result
'            End If
'        Next
'        txtbox.text = finalresult
'        If oldselstart > Len(txtbox.text) Then
'            txtbox.selstart = Len(txtbox.text)
'        Else
'            txtbox.selstart = oldselstart
'        End If
'    End If
'
'    If val(txtbox.text) < lowerBound Then
'        txtbox.text = lowerBound
'    End If
'
'    If val(txtbox.text) > upperBound Then
'        txtbox.text = upperBound
'    End If
'
'End Sub

Public Property Get ListCount() As Integer
10        ListCount = propcount
End Property

Public Property Get ListIndex() As Integer
10        If curlstItem Is Nothing Then
20            ListIndex = -1
30        Else
40            ListIndex = curlstItem.Index - 1
50        End If
End Property

Public Property Get Enabled() As Boolean
10        Enabled = lst.Enabled
End Property

Public Property Let Enabled(ByVal newval As Boolean)
10        If lst.Enabled = False And newval = True Then
20            lst.HideSelection = False
30            lst.ColumnHeaders(1).Text = "Property"
40            lst.ColumnHeaders(2).Text = "Value"
50            UpdateList
60        ElseIf lst.Enabled = True And newval = False Then
70            lst.HideSelection = True
80            lst.ColumnHeaders(1).Text = ""
90            lst.ColumnHeaders(2).Text = ""
100           lst.ListItems.Clear
110       End If
          
120       lst.Enabled = newval
End Property



Private Sub removeInvalidExpressionCharacters(ByRef txtbox As TextBox, lowerBound As Single, upperBound As Single, Optional dec As Boolean = False)
    
    
'    Dim i As Integer
'    For i = 1 To Len(txtbox.test)
'        if
    
    
'    If (Not IsNumeric(txtbox.text) And (txtbox.text <> "-" Or lowerBound >= 0)) _
'        Or InStr(txtbox.text, "e") > 0 Or InStr(txtbox.text, "E") > 0 _
'        Or (Not dec And (InStr(txtbox.text, ".") > 0 Or InStr(txtbox.text, ",") > 0)) _
'        Or (lowerBound < 0 And InStr(2, txtbox.text, "-") > 1) _
'        Or (lowerBound >= 0 And InStr(txtbox.text, "-") > 0) Then
'
'        Dim oldselstart As Integer
'        oldselstart = txtbox.selstart - 1    'char  typed so always one more
'        If oldselstart < 0 Then oldselstart = 0
'
'        'remove all characters aside from nrs
'
'        Dim finalresult As String
'        For i = 1 To Len(txtbox.text)
'            If (Asc(Mid$(txtbox.text, i, 1)) < Asc("0") Or _
'                Asc(Mid$(txtbox.text, i, 1)) > Asc("9")) Then
'                Dim result As String
'                If i - 1 >= 1 Then result = Mid$(txtbox.text, 1, i - 1)
'                If i + 1 <= Len(txtbox.text) Then result = result + Mid$(txtbox.text, i + 1, Len(txtbox.text) - (i))
'                finalresult = result
'            End If
'        Next
'        txtbox.text = finalresult
'        If oldselstart > Len(txtbox.text) Then
'            txtbox.selstart = Len(txtbox.text)
'        Else
'            txtbox.selstart = oldselstart
'        End If
'    End If
'
'    If val(txtbox.text) < lowerBound Then
'        txtbox.text = lowerBound
'    End If
'
'    If val(txtbox.text) > upperBound Then
'        txtbox.text = upperBound
'    End If

End Sub

