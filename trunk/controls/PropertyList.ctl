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

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
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
    
    Set overlstItem = lst.HitTest(X, Y)
    
    If overlstItem Is Nothing Then
        lst.tooltiptext = ""
    Else
        lst.tooltiptext = props(overlstItem.Index - 1).tooltip
    End If
End Sub

Private Sub lst_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If edit Then
        EditOff (True)
        Exit Sub
    End If
    
    curX = X
    curY = Y
    Call EditOn
End Sub

Private Sub lstChoice_Click()
    If edit Then EditOff (True)
End Sub

Private Sub txt_Change()
    If edit And props(curlstItem.Index - 1).Type = p_number Then
        Call removeDisallowedCharacters(txt, CSng(props(curlstItem.Index - 1).lbnd), CSng(props(curlstItem.Index - 1).ubnd), False)
    End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then EditOff (True)
    If KeyAscii = vbKeyEscape Then EditOff (False)
        
End Sub

Private Sub UserControl_Initialize()
    ReDim props(0)
    propcount = 0
End Sub

Private Sub UserControl_LostFocus()
    If edit Then EditOff (True)
End Sub

Private Sub UserControl_Resize()
    lst.width = UserControl.width
    lst.height = UserControl.height
    
    If propcount > 0 Then
        If (propcount + 1) * lst.height > UserControl.height Then
        'scrollbar will appear
            lst.ColumnHeaders(1).width = (lst.width / 2) - 9 * Screen.TwipsPerPixelX
            lst.ColumnHeaders(2).width = lst.width / 2 - 9 * Screen.TwipsPerPixelX
        Else
            lst.ColumnHeaders(1).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
            lst.ColumnHeaders(2).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
        End If
    Else
        lst.ColumnHeaders(1).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
        lst.ColumnHeaders(2).width = lst.width / 2 - 1 * Screen.TwipsPerPixelX
    End If
End Sub

Sub AddProperty(name As String, proptype As propertyType, Optional lbnd As Long = -2147483647, Optional ubnd As Long = 2147483647)
    If propcount > UBound(props) Then
        ReDim Preserve props(UBound(props) + 5)
    End If
    
    Dim val As Integer
    val = getPropertyIndex(name)
    If val <> -1 Then
        Call Err.Raise(10002, , "Property '" & name & "' already exists")
        Exit Sub
    End If
    
    props(propcount).name = name
    props(propcount).Type = proptype
    
    If proptype = p_number Then
        props(propcount).value = 0
        props(propcount).lbnd = lbnd
        props(propcount).ubnd = ubnd
    ElseIf proptype = p_list Then
        ReDim props(propcount).choices(0)
    End If
    
    If lst.Enabled Then Call lst.ListItems.add(propcount + 1, name, name)
    
    If (propcount + 1) * lst.ListItems(name).height > UserControl.height Then
        'scrollbar will appear
        lst.ColumnHeaders(1).width = (lst.width / 2) - 9 * Screen.TwipsPerPixelX
        lst.ColumnHeaders(2).width = lst.width / 2 - 9 * Screen.TwipsPerPixelX
    End If
            
    
    propcount = propcount + 1
End Sub

Sub setPropertyToolTipText(name As String, tooltiptext As String)
    Dim idx As Integer
    idx = getPropertyIndex(name)
    
    If idx = -1 Then
        Call Err.Raise(10000, , "Property '" & name & "' does not exist")
        Exit Sub
    End If
    
    props(idx).tooltip = tooltiptext
End Sub

Sub setPropertyChoiceList(name As String, chlst() As String)
    Dim idx As Integer
    idx = getPropertyIndex(name)
        
    If idx = -1 Then
        Call Err.Raise(10000, , "Property '" & name & "' does not exist")
        Exit Sub
    End If
    
    If props(idx).Type = p_list Then
        props(idx).choices = chlst
        
        props(idx).value = 0
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
    Else
        Call Err.Raise(10001, , "Property is not a list type")
    End If
End Sub

Sub setPropertyNumberBoundaries(name As String, lbnd As Long, ubnd As Long)
    Dim idx As Integer
    idx = getPropertyIndex(name)
    
    If idx = -1 Then
        Call Err.Raise(10000, , "Property '" & name & "' does not exist")
        Exit Sub
    End If
    
    If props(idx).Type = p_number Then
        props(idx).lbnd = lbnd
        props(idx).ubnd = ubnd
    Else
        Call Err.Raise(10001, "Property is not a number type")
    End If
End Sub

Sub setPropertyValue(name As String, value As String)
    
    Dim idx As Integer
    idx = getPropertyIndex(name)
    
    If idx = -1 Then
        Call Err.Raise(10000, , "Property '" & name & "' does not exist")
        Exit Sub
    End If
    
    If props(idx).Type = p_list Then
        props(idx).value = val(value)
    End If
    
    If lst.Enabled Then
        If props(idx).Type = p_list Then
            lst.ListItems(props(idx).name).SubItems(1) = props(idx).choices(value)
        Else
            lst.ListItems(props(idx).name).SubItems(1) = value
        End If
    End If
End Sub

Sub setPropertyLocked(name As String, locked As Boolean)
    Dim idx As Integer
    idx = getPropertyIndex(name)
    
    If idx = -1 Then
        Call Err.Raise(10000, , "Property '" & name & "' does not exist")
        Exit Sub
    End If
    
    props(idx).locked = locked
End Sub

Function getPropertyValue(name As String) As String
    Dim idx As Integer
    idx = getPropertyIndex(name)
    
    If idx = -1 Then
        Call Err.Raise(10000, , "Property '" & name & "' does not exist")
        Exit Function
    End If
    
    getPropertyValue = props(idx).value
End Function

Sub RemoveProperty(name As String)
    If propcount < UBound(props) - 5 Then
        ReDim Preserve props(UBound(props) - 5)
    End If
    
    Dim idx As Integer
    idx = getPropertyIndex(name)
    
    If idx <> -1 Then
        Dim i As Integer
        For i = idx + 1 To propcount - 1
            props(i - 1) = props(i)
        Next
    End If

End Sub

Private Function getPropertyIndex(name As String) As Integer
    Dim i As Integer
    For i = 0 To propcount - 1
        If props(i).name = name Then
            getPropertyIndex = i
            Exit Function
        End If
    Next
    
    getPropertyIndex = -1
End Function

Sub EditOn()
    Set curlstItem = lst.HitTest(curX, curY)
    
    If curlstItem Is Nothing Then
        Exit Sub
    End If
    
    If props(curlstItem.Index - 1).locked Then
        Exit Sub
    End If
    
    Select Case props(curlstItem.Index - 1).Type
    
        Case p_text, p_number
            txt.width = lst.ColumnHeaders(2).width - 2 * Screen.TwipsPerPixelX
            txt.height = curlstItem.height - 1 * Screen.TwipsPerPixelY
            txt.Left = lst.Left + curlstItem.Left + lst.ColumnHeaders(1).width
            txt.Top = lst.Top + curlstItem.Top + 2 * Screen.TwipsPerPixelY
            txt.Text = curlstItem.SubItems(1)
            txt.selstart = 0
            txt.sellength = Len(txt.Text)
            txt.visible = True
            txt.setfocus
            
        Case p_list

            lstChoice.Left = lst.Left + curlstItem.Left + lst.ColumnHeaders(1).width
            lstChoice.Top = lst.Top + curlstItem.Top + 1 * Screen.TwipsPerPixelY + curlstItem.height
            lstChoice.width = lst.ColumnHeaders(2).width - 1 * Screen.TwipsPerPixelX

            
            lstChoice.Clear
            Dim i As Integer
            
            For i = 0 To UBound(props(curlstItem.Index - 1).choices)
                If props(curlstItem.Index - 1).choices(i) <> "" Then
                    lstChoice.addItem props(curlstItem.Index - 1).choices(i)
                End If
            '    If lstChoice.list(lstChoice.Listcount - 1) = props(curlstItem.Index - 1).value Then
           '         lstChoice.ListIndex = i
           '     End If
            Next
            lstChoice.ListIndex = val(props(curlstItem.Index - 1).value)
            
            If lstChoice.ListIndex = -1 Then lstChoice.ListIndex = 0

            If lstChoice.ListCount < 5 Then
                lstChoice.height = (curlstItem.height + 1) * lstChoice.ListCount
            Else
                lstChoice.height = curlstItem.height * 5
            End If
            
            
            'show the lstchoice, give it to the desktopwindow, and move it to the right position
            Dim pt As POINTAPI, parentrect As RECT
            
            pt.X = ScaleX(lstChoice.Left, vbTwips, vbPixels) '+ UserControl.parent.Left
            pt.Y = ScaleY(lstChoice.Top, vbTwips, vbPixels)

            ClientToScreen UserControl.hWnd, pt
          
          
            GetWindowRect UserControl.ContainerHwnd, parentrect
            
            setParent lstChoice.hWnd, UserControl.ContainerHwnd
          
          'Make sure the box appears within the parent container
          If pt.Y + ScaleY(lstChoice.height, vbTwips, vbPixels) > parentrect.Bottom Then
              pt.Y = parentrect.Bottom - ScaleY(lstChoice.height, vbTwips, vbPixels)
          End If
                          
          lstChoice.Move ScaleX(pt.X - parentrect.Left, vbPixels, vbTwips), ScaleY(pt.Y - parentrect.Top, vbPixels, vbTwips)


          lstChoice.visible = True
            ShowWindow lstChoice.hWnd, SW_SHOWNORMAL


            
    End Select
    
    edit = True
End Sub

Sub EditOff(apply As Boolean)
    Dim changed As Boolean
    
    If Not edit Then Exit Sub
    
    Select Case props(curlstItem.Index - 1).Type
        Case p_text, p_number
                        
            If apply Then
                If props(curlstItem.Index - 1).value <> txt.Text Then changed = True
                props(curlstItem.Index - 1).value = txt.Text
            End If
                
            curlstItem.SubItems(1) = txt.Text
            txt.visible = False
            If changed Then RaiseEvent PropertyChanged(props(curlstItem.Index - 1).name)
            
        Case p_list
            If apply Then
                If props(curlstItem.Index - 1).value <> lstChoice.ListIndex Then changed = True
                props(curlstItem.Index - 1).value = lstChoice.ListIndex
            End If
            curlstItem.SubItems(1) = lstChoice.list(lstChoice.ListIndex)
            lstChoice.visible = False
            'show the lstchoice, give it to the desktopwindow, and move it to the right position
            setParent lstChoice.hWnd, UserControl.hWnd
            lstChoice.Left = 0
            lstChoice.Top = 0
            'ShowWindow lstChoice.hWnd, SW_SHOWNOACTIVATE
            
            
            If changed Then RaiseEvent PropertyChanged(props(curlstItem.Index - 1).name)
            
    End Select
    edit = False
    
End Sub

Sub UpdateList()
    lst.ListItems.Clear
    
    Dim i As Integer
    For i = 0 To propcount - 1
        Call lst.ListItems.add(, props(i).name, props(i).name)
        If props(i).Type = p_list Then
            lst.ListItems.item(props(i).name).SubItems(1) = props(i).choices(props(i).value)
        Else
            lst.ListItems.item(props(i).name).SubItems(1) = props(i).value
        End If
    Next
    
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
    ListCount = propcount
End Property

Public Property Get ListIndex() As Integer
    If curlstItem Is Nothing Then
        ListIndex = -1
    Else
        ListIndex = curlstItem.Index - 1
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = lst.Enabled
End Property

Public Property Let Enabled(ByVal newval As Boolean)
    If lst.Enabled = False And newval = True Then
        lst.HideSelection = False
        lst.ColumnHeaders(1).Text = "Property"
        lst.ColumnHeaders(2).Text = "Value"
        UpdateList
    ElseIf lst.Enabled = True And newval = False Then
        lst.HideSelection = True
        lst.ColumnHeaders(1).Text = ""
        lst.ColumnHeaders(2).Text = ""
        lst.ListItems.Clear
    End If
    
    lst.Enabled = newval
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

Private Sub UserControl_Terminate()
    '...
End Sub
