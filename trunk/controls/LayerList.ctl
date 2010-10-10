VERSION 5.00
Begin VB.UserControl LayerList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   MaskColor       =   &H0000FF00&
   ScaleHeight     =   2670
   ScaleWidth      =   3510
   ToolboxBitmap   =   "LayerList.ctx":0000
   Begin VB.PictureBox picDropDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00808080&
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   3465
      TabIndex        =   5
      Top             =   255
      Width           =   3495
      Begin VB.VScrollBar scroll 
         Height          =   2415
         LargeChange     =   5
         Left            =   3255
         Max             =   0
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox piccolor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1200
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHighlight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   1440
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdDummy 
      Caption         =   "dummy"
      Height          =   195
      Left            =   -750
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdDropDown 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3270
      Picture         =   "LayerList.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   240
      Picture         =   "LayerList.ctx":0504
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "LayerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function setParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Const SW_SHOWNOACTIVATE = 4

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Private Type WINDOWPLACEMENT
'    Length As Long
'    flags As Long
'    showCmd As Long
'    ptMinPosition As POINTAPI
'    ptMaxPosition As POINTAPI
'    rcNormalPosition As RECT
'End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type item
    visible As Boolean
    name As String
    color As Long
End Type

Dim items() As item
Dim itemcount As Integer

Dim drawFull As Boolean

Dim c_listindex As Integer

Private Const ICON_VISIBLE = 0
Private Const ICON_INVISIBLE = 16
Private Const ICON_DELETE = 32
Private Const ICON_EDIT = 48

Dim bc As BLENDFUNCTION
Dim lbc As Long

'Dim ancestor As Long

Dim Amount_Shown As Integer


Dim hOldParent As Long


Event AddItemClick()
Event Change()
Event DeleteItemClick(Index As Integer)
Event EditItemClick(Index As Integer)
Event VisibiltyChanged(Index As Integer)
Event ChangeItemColor(Index As Integer)
Event RightClick(Index As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub cmdDropDown_Click()
    If drawFull Then
        HideList
    Else
        ShowList
    End If

    On Error Resume Next
    cmdDummy.setfocus
End Sub

Public Property Let ArrowTooltipText(n_Tooltiptext As String)
    cmdDropDown.tooltiptext = n_Tooltiptext
End Property

Public Property Let DropdownTooltipText(n_Tooltiptext As String)
    picDropDown.tooltiptext = n_Tooltiptext
End Property

Private Sub picDropDown_LostFocus()
    HideList
End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calculate the clicked item
    Dim curItem As Integer
    curItem = (Y / Screen.TwipsPerPixelY + 16 - 2) \ 16 + scroll.value - 1    '((((Y) Mod 16) + scroll.Value

    If curItem >= itemcount Then
        'also check if user did not click on a disabled scrollbar
        If X < scroll.Left Then
            RaiseEvent AddItemClick
            HideList
            Exit Sub
        Else
            Exit Sub
        End If
    ElseIf curItem = -1 Then
        curItem = ListIndex
    End If

    Dim Xpx As Integer
    Xpx = X / Screen.TwipsPerPixelX

    If Xpx \ 16 = 0 Then
        'change color
        RaiseEvent ChangeItemColor(curItem)
    ElseIf Xpx \ 16 = 1 Then
        'delete
        RaiseEvent DeleteItemClick(curItem)
    ElseIf Xpx \ 16 = 2 Then
        RaiseEvent EditItemClick(curItem)
    ElseIf Xpx \ 16 = 3 Then
        items(curItem).visible = Not items(curItem).visible
        Redraw
        RaiseEvent VisibiltyChanged(curItem)
    ElseIf Button = vbRightButton Then
        'right-clicked on name
        RaiseEvent RightClick(curItem)
    Else
        'clicked on name
        If curItem = -1 Then
            If drawFull Then
                Call HideList
            Else
                Call ShowList
            End If
        Else
            ListIndex = curItem
            HideList
            RaiseEvent Change
        End If
    End If

    frmGeneral.Label6.Caption = "ListIndex=" & ListIndex
End Sub

Private Sub scroll_Change()
    Redraw
End Sub

Private Sub scroll_GotFocus()
    cmdDummy.setfocus
End Sub

Private Sub scroll_Scroll()
    Redraw
End Sub

Private Sub UserControl_Initialize()
    ReDim items(0)
    itemcount = 0
    ListIndex = -1
    bc.SourceConstantAlpha = 128
    RtlMoveMemory lbc, bc, 4
    Amount_Shown = 5
    
    hOldParent = 0
End Sub

Private Sub UserControl_ExitFocus()
    HideList
End Sub

Private Sub UserControl_LostFocus()
    HideList
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    cmdDropDown.Left = UserControl.ScaleWidth - cmdDropDown.width
    scroll.Left = UserControl.ScaleWidth - scroll.width

    If drawFull Then
        UserControl.height = ((17 + 3) * Screen.TwipsPerPixelY)
    Else
        UserControl.height = ((16 + 3) * Screen.TwipsPerPixelY)
    End If

    If UserControl.width < 2000 Then
        UserControl.width = 2000
    End If

    picHighlight.width = UserControl.width

    Redraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calculate the clicked item
    Dim curItem As Integer
    curItem = (Y / Screen.TwipsPerPixelY - 2) \ 16 + scroll.value - 1    '((((Y) Mod 16) + scroll.Value

    If curItem >= itemcount Then
        RaiseEvent AddItemClick
        HideList
        Exit Sub
    ElseIf curItem = -1 Then
        curItem = ListIndex
    End If

    If itemcount = 0 Then
        'there are no elements
        RaiseEvent AddItemClick
        If drawFull Then HideList
        Exit Sub
    End If

    Dim Xpx As Integer
    Xpx = X / Screen.TwipsPerPixelX

    If Xpx \ 16 = 0 Then
        'change color
        RaiseEvent ChangeItemColor(curItem)
    ElseIf Xpx \ 16 = 1 Then
        'delete
        RaiseEvent DeleteItemClick(curItem)
    ElseIf Xpx \ 16 = 2 Then
        RaiseEvent EditItemClick(curItem)
    ElseIf Xpx \ 16 = 3 Then
        items(curItem).visible = Not items(curItem).visible
        Redraw
        RaiseEvent VisibiltyChanged(curItem)
    ElseIf Button = vbRightButton And itemcount > 0 Then
        'right-clicked on name
        RaiseEvent RightClick(curItem)
    Else
        'clicked on name
        If Y / Screen.TwipsPerPixelY <= 16 And itemcount > 0 Then
            If drawFull Then
                Call HideList
            Else
                Call ShowList
            End If
        ElseIf itemcount = 0 Then
            'there are no elements
            RaiseEvent AddItemClick
            HideList
        Else
            ListIndex = curItem
            HideList
            RaiseEvent Change

        End If
    End If
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    If cmdDummy.visible Then cmdDummy.setfocus
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Amount_Shown = PropBag.ReadProperty("LinesShown", 5)
End Sub

Private Sub UserControl_Terminate()
    If hOldParent <> 0 Then
        setParent picDropDown.hWnd, hOldParent
        hOldParent = 0
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("LinesShown", Amount_Shown, 5)
End Sub

Sub addItem(name As String, visible As Boolean, color As Long)
    If itemcount >= UBound(items) Then
        ReDim Preserve items(UBound(items) + 5)
    End If

    items(itemcount).name = name
    items(itemcount).visible = visible
    items(itemcount).color = color

    itemcount = itemcount + 1

    cmdDropDown.Enabled = True
    If itemcount = 1 Then
        'we added 1, auto-set selected on first element
        ListIndex = 0
    End If

    Redraw
End Sub

Sub removeItem(Index As Integer)
    Dim i As Integer
    For i = Index + 1 To itemcount - 1
        items(i - 1) = items(i)
    Next

    itemcount = itemcount - 1

    If itemcount = 0 Then
        cmdDropDown.Enabled = False
        If drawFull Then HideList
    End If

    Redraw
End Sub

Sub ShowList()
    drawFull = True

    'change the scroll max to the amount_shown
    If itemcount > Amount_Shown Then
        scroll.Min = 0
        scroll.Max = itemcount - Amount_Shown + 1
        scroll.Enabled = True
      If ListIndex > Amount_Shown Then
          If ListIndex < scroll.Max Then
              scroll.value = ListIndex
          Else
              scroll.value = scroll.Max
          End If
      Else
        scroll.value = 0
      End If
      scroll.LargeChange = Amount_Shown
    Else
        scroll.Max = 0
        scroll.Enabled = False
    End If

    'resize the usercontrol
    Call UserControl_Resize

    ' calculate how many are drawn
    Dim Amount_Drawn As Integer
    If itemcount > Amount_Shown Then
        Amount_Drawn = Amount_Shown
        picDropDown.height = (Amount_Drawn * 16 + 2) * Screen.TwipsPerPixelY
    Else
        Amount_Drawn = itemcount
        picDropDown.height = (Amount_Drawn * 16 + 16 + 1 + 2) * Screen.TwipsPerPixelY
    End If

    'change the controls size to fit
    scroll.height = picDropDown.height - 1 * Screen.TwipsPerPixelY
    picDropDown.width = UserControl.width - 3 * Screen.TwipsPerPixelX
    scroll.Left = picDropDown.width - scroll.width

    'show the picdropdown, give it to the desktopwindow, and move it to the right position
    Dim pt As POINTAPI
    pt.X = -1
    pt.Y = 17
    ClientToScreen UserControl.hWnd, pt

    hOldParent = setParent(picDropDown.hWnd, GetDesktopWindow)

    picDropDown.Move ScaleX(pt.X, vbPixels, vbTwips), ScaleY(pt.Y, vbPixels, vbTwips)
    ShowWindow picDropDown.hWnd, SW_SHOWNOACTIVATE
    picDropDown.visible = True

    Redraw
End Sub

Sub HideList()
    drawFull = False

    If hOldParent <> 0 Then
  setParent picDropDown.hWnd, hOldParent
  hOldParent = 0
    End If
    Call UserControl_Resize
    picDropDown.visible = False

    Redraw
End Sub



Sub Redraw()


    UserControl.Cls
    picDropDown.Cls

    Dim yoffset As Integer

    'draw selected (on usercontrol)
    If itemcount <> 0 Then
        piccolor.BackColor = items(ListIndex).color
        BitBlt UserControl.hDC, 0, yoffset, 16, 16, piccolor.hDC, 0, 0, vbSrcCopy

        BitBlt UserControl.hDC, 16, yoffset, 16, 16, picicons.hDC, ICON_DELETE, 0, vbSrcCopy
        BitBlt UserControl.hDC, 32, yoffset, 16, 16, picicons.hDC, ICON_EDIT, 0, vbSrcCopy
        BitBlt UserControl.hDC, 48, yoffset, 16, 16, picicons.hDC, IIf(items(ListIndex).visible, ICON_VISIBLE, ICON_INVISIBLE), 0, vbSrcCopy


        UserControl.CurrentX = 65 * Screen.TwipsPerPixelX
        UserControl.CurrentY = (yoffset + IIf(TextHeight(items(ListIndex).name) / Screen.TwipsPerPixelY < 16, (16 - TextHeight(items(ListIndex).name) / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
        UserControl.ForeColor = vbBlack
        UserControl.Print items(ListIndex).name

        'draw a line to seperate from the drop down (not visible if not dropped down)
        UserControl.Line (0, 16 * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth, 16 * Screen.TwipsPerPixelY), &HC0C0C0

    End If


    If itemcount <> 0 And drawFull Then

        Dim Amount_Drawn As Integer
        If itemcount > Amount_Shown Then
            Amount_Drawn = Amount_Shown
        Else
            Amount_Drawn = itemcount
        End If

        Dim j As Integer
        For j = scroll.value To scroll.value + Amount_Drawn - 1
            If j < itemcount Then

                'draw icons
                BitBlt picDropDown.hDC, 16, yoffset + (j - scroll.value) * 16, 16, 16, picicons.hDC, ICON_DELETE, 0, vbSrcCopy
                BitBlt picDropDown.hDC, 32, yoffset + (j - scroll.value) * 16, 16, 16, picicons.hDC, ICON_EDIT, 0, vbSrcCopy
                BitBlt picDropDown.hDC, 48, yoffset + (j - scroll.value) * 16, 16, 16, picicons.hDC, IIf(items(j).visible, ICON_VISIBLE, ICON_INVISIBLE), 0, vbSrcCopy

                'draw the color of the region
                piccolor.BackColor = items(j).color
                BitBlt picDropDown.hDC, 0, yoffset + (j - scroll.value) * 16, 16, 16, piccolor.hDC, 0, 0, vbSrcCopy

                'if item is selected, highlight it
                If j = ListIndex Then
                    Call AlphaBlend(picDropDown.hDC, 0, yoffset + (j - scroll.value) * 16, picHighlight.ScaleWidth, picHighlight.ScaleHeight, picHighlight.hDC, 0, 0, picHighlight.ScaleWidth, picHighlight.ScaleHeight, lbc)
                End If

                ' print the name
                picDropDown.CurrentX = 65 * Screen.TwipsPerPixelX
                picDropDown.CurrentY = (yoffset + (j - scroll.value) * 16 + IIf(TextHeight(items(j).name) / Screen.TwipsPerPixelY < 16, (16 - TextHeight(items(j).name) / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
                picDropDown.ForeColor = vbBlack
                picDropDown.Print items(j).name
            End If
        Next

        ' change the yoffset depending
        If itemcount > Amount_Shown Then
            yoffset = (Amount_Drawn - 1) * 16
        Else
            yoffset = (Amount_Drawn - 1) * 16 + 16
        End If
    End If

    If itemcount = 0 And Not drawFull Then
        'draw the Add New Region on the usercontrol because there are no regions
        UserControl.CurrentX = (((UserControl.ScaleWidth - scroll.width) / Screen.TwipsPerPixelX) / 2 - (TextWidth("Add New Region...") / Screen.TwipsPerPixelX) / 2) * Screen.TwipsPerPixelX
        UserControl.CurrentY = (1 + yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
        UserControl.ForeColor = vbBlack
        UserControl.Print "Add New Region..."

    ElseIf drawFull And scroll.value = scroll.Max Then
        picDropDown.Line (0, (yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth, (yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY), &HC0C0C0

        'draw the Add New Region
        picDropDown.CurrentX = (((UserControl.ScaleWidth - scroll.width) / Screen.TwipsPerPixelX) / 2 - (TextWidth("Add New Region...") / Screen.TwipsPerPixelX) / 2) * Screen.TwipsPerPixelX
        picDropDown.CurrentY = (1 + yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
        picDropDown.ForeColor = vbBlack
        picDropDown.Print "Add New Region..."
    End If


End Sub
Sub Clear()
    Erase items
    ReDim items(0)
    itemcount = 0
    c_listindex = -1
    cmdDropDown.Enabled = False

    Redraw
End Sub

Sub Sort()
    If itemcount = 0 Then Exit Sub

    Dim selitem As item
    selitem = items(ListIndex)

    Call Quicksort(items, 0, itemcount - 1)

    'retrieve the index of the current selected item
    Dim i As Integer
    For i = 0 To itemcount - 1
        If selitem.name = items(i).name And selitem.color = items(i).color And selitem.visible = items(i).visible Then
            ListIndex = i
            Exit For
        End If
    Next

    Redraw
End Sub

' Use Quicksort to sort a list of strings.
'
' This code is from the book "Ready-to-Run
' Visual Basic Algorithms" by Rod Stephens.
' http://www.vb-helper.com/vba.htm
Private Sub Quicksort(list() As item, ByVal Min As Long, _
                      ByVal Max As Long)
    Dim mid_value As item
    Dim hi As Long
    Dim lo As Long
    Dim i As Long

    ' If there is 0 or 1 item in the list,
    ' this sublist is sorted.
    If Min >= Max Then Exit Sub

    ' Pick a dividing value.
    i = Int((Max - Min + 1) * Rnd + Min)
    mid_value = list(i)

    ' Swap the dividing value to the front.
    list(i) = list(Min)

    lo = Min
    hi = Max
    Do
        ' Look down from hi for a value < mid_value.
        Do While list(hi).name >= mid_value.name
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            list(lo) = mid_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        list(lo) = list(hi)

        ' Look up from lo for a value >= mid_value.
        lo = lo + 1
        Do While list(lo).name < mid_value.name
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            list(hi) = mid_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        list(hi) = list(lo)
    Loop

    ' Sort the two sublists.
    Quicksort list, Min, lo - 1
    Quicksort list, lo + 1, Max
End Sub


Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = c_listindex
End Property

Public Property Let ListIndex(ByVal lstindex As Integer)
    Dim oldidx As Integer
    oldidx = c_listindex
    
    If itemcount = 0 Then
        c_listindex = -1
    Else
        If lstindex < 0 Or lstindex >= itemcount Then
            Call Err.Raise(10001, , "The index '" & lstindex & "' is outside the boundaries of the list")
            Exit Property
        End If

        c_listindex = lstindex
    End If
    If oldidx <> c_listindex Then
  RaiseEvent Change
    End If
End Property

Public Property Get ListCount() As Integer
    ListCount = itemcount
End Property

Public Property Get LinesShown() As Integer
Attribute LinesShown.VB_Description = "Maximum number of lines that are shown when the list is dropped down"
    LinesShown = Amount_Shown
End Property

Public Property Let LinesShown(ByVal newval As Integer)
    Amount_Shown = newval
    PropertyChanged "LinesShown"

End Property

