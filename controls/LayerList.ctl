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

Event AddItemClick()
Event change()
Event DeleteItemClick(Index As Integer)
Event EditItemClick(Index As Integer)
Event VisibiltyChanged(Index As Integer)
Event ChangeItemColor(Index As Integer)
Event RightClick(Index As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub cmdDropDown_Click()
10        If drawFull Then
20            HideList
30        Else
40            ShowList
50        End If

60        On Error Resume Next
70        cmdDummy.setfocus
End Sub

Public Property Let ArrowTooltipText(n_Tooltiptext As String)
10        cmdDropDown.tooltiptext = n_Tooltiptext
End Property

Public Property Let DropdownTooltipText(n_Tooltiptext As String)
10        picDropDown.tooltiptext = n_Tooltiptext
End Property

Private Sub picDropDown_LostFocus()
10        HideList
End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'calculate the clicked item
          Dim curItem As Integer
10        curItem = (Y / Screen.TwipsPerPixelY + 16 - 2) \ 16 + scroll.value - 1    '((((Y) Mod 16) + scroll.Value

20        If curItem >= itemcount Then
              'also check if user did not click on a disabled scrollbar
30            If X < scroll.Left Then
40                RaiseEvent AddItemClick
50                HideList
60                Exit Sub
70            Else
80                Exit Sub
90            End If
100       ElseIf curItem = -1 Then
110           curItem = ListIndex
120       End If

          Dim Xpx As Integer
130       Xpx = X / Screen.TwipsPerPixelX

140       If Xpx \ 16 = 0 Then
              'change color
150           RaiseEvent ChangeItemColor(curItem)
160       ElseIf Xpx \ 16 = 1 Then
              'delete
170           RaiseEvent DeleteItemClick(curItem)
180       ElseIf Xpx \ 16 = 2 Then
190           RaiseEvent EditItemClick(curItem)
200       ElseIf Xpx \ 16 = 3 Then
210           items(curItem).visible = Not items(curItem).visible
220           Redraw
230           RaiseEvent VisibiltyChanged(curItem)
240       ElseIf Button = vbRightButton Then
              'right-clicked on name
250           RaiseEvent RightClick(curItem)
260       Else
              'clicked on name
270           If curItem = -1 Then
280               If drawFull Then
290                   Call HideList
300               Else
310                   Call ShowList
320               End If
330           Else
340               ListIndex = curItem
350               HideList
360               RaiseEvent change
370           End If
380       End If

End Sub

Private Sub scroll_Change()
10        Redraw
End Sub

Private Sub scroll_GotFocus()
10        cmdDummy.setfocus
End Sub

Private Sub scroll_Scroll()
10        Redraw
End Sub

Private Sub UserControl_Initialize()
10        ReDim items(0)
20        itemcount = 0
30        ListIndex = -1
40        bc.SourceConstantAlpha = 128
50        RtlMoveMemory lbc, bc, 4
60        Amount_Shown = 5

End Sub

Private Sub UserControl_ExitFocus()
10        HideList
End Sub

Private Sub UserControl_LostFocus()
10        HideList
End Sub

Private Sub UserControl_Resize()
10        On Error Resume Next
20        cmdDropDown.Left = UserControl.ScaleWidth - cmdDropDown.width
30        scroll.Left = UserControl.ScaleWidth - scroll.width

40        If drawFull Then
50            UserControl.height = ((17 + 3) * Screen.TwipsPerPixelY)
60        Else
70            UserControl.height = ((16 + 3) * Screen.TwipsPerPixelY)
80        End If

90        If UserControl.width < 2000 Then
100           UserControl.width = 2000
110       End If

120       picHighlight.width = UserControl.width

130       Redraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'calculate the clicked item
          Dim curItem As Integer
10        curItem = (Y / Screen.TwipsPerPixelY - 2) \ 16 + scroll.value - 1    '((((Y) Mod 16) + scroll.Value

20        If curItem >= itemcount Then
30            RaiseEvent AddItemClick
40            HideList
50            Exit Sub
60        ElseIf curItem = -1 Then
70            curItem = ListIndex
80        End If

90        If itemcount = 0 Then
              'there are no elements
100           RaiseEvent AddItemClick
110           If drawFull Then HideList
120           Exit Sub
130       End If

          Dim Xpx As Integer
140       Xpx = X / Screen.TwipsPerPixelX

150       If Xpx \ 16 = 0 Then
              'change color
160           RaiseEvent ChangeItemColor(curItem)
170       ElseIf Xpx \ 16 = 1 Then
              'delete
180           RaiseEvent DeleteItemClick(curItem)
190       ElseIf Xpx \ 16 = 2 Then
200           RaiseEvent EditItemClick(curItem)
210       ElseIf Xpx \ 16 = 3 Then
220           items(curItem).visible = Not items(curItem).visible
230           Redraw
240           RaiseEvent VisibiltyChanged(curItem)
250       ElseIf Button = vbRightButton And itemcount > 0 Then
              'right-clicked on name
260           RaiseEvent RightClick(curItem)
270       Else
              'clicked on name
280           If Y / Screen.TwipsPerPixelY <= 16 And itemcount > 0 Then
290               If drawFull Then
300                   Call HideList
310               Else
320                   Call ShowList
330               End If
340           ElseIf itemcount = 0 Then
                  'there are no elements
350               RaiseEvent AddItemClick
360               HideList
370           Else
380               ListIndex = curItem
390               HideList
400               RaiseEvent change

410           End If
420       End If
End Sub

Private Sub UserControl_Show()
10        On Error Resume Next
20        If cmdDummy.visible Then cmdDummy.setfocus
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
10        Amount_Shown = PropBag.ReadProperty("LinesShown", 5)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
10        Call PropBag.WriteProperty("LinesShown", Amount_Shown, 5)
End Sub

Sub addItem(name As String, visible As Boolean, color As Long)
10        If itemcount >= UBound(items) Then
20            ReDim Preserve items(UBound(items) + 5)
30        End If

40        items(itemcount).name = name
50        items(itemcount).visible = visible
60        items(itemcount).color = color

70        itemcount = itemcount + 1

80        cmdDropDown.Enabled = True
90        If itemcount = 1 Then
              'we added 1, auto-set selected on first element
100           ListIndex = 0
110       End If

120       Redraw
End Sub

Sub removeItem(Index As Integer)
          Dim i As Integer
10        For i = Index + 1 To itemcount - 1
20            items(i - 1) = items(i)
30        Next

40        itemcount = itemcount - 1

50        If itemcount = 0 Then
60            cmdDropDown.Enabled = False
70            If drawFull Then HideList
80        End If

90        Redraw
End Sub

Sub ShowList()
10        drawFull = True

          'change the scroll max to the amount_shown
20        If itemcount > Amount_Shown Then
30            scroll.Min = 0
40            scroll.Max = itemcount - Amount_Shown + 1
50            scroll.Enabled = True
            If ListIndex > Amount_Shown Then
                If ListIndex < scroll.Max Then
                    scroll.value = ListIndex
                Else
                    scroll.value = scroll.Max
                End If
            Else
60            scroll.value = 0
            End If
            scroll.LargeChange = Amount_Shown
70        Else
80            scroll.Max = 0
90            scroll.Enabled = False
100       End If

          'resize the usercontrol
110       Call UserControl_Resize

          ' calculate how many are drawn
          Dim Amount_Drawn As Integer
120       If itemcount > Amount_Shown Then
130           Amount_Drawn = Amount_Shown
140           picDropDown.height = (Amount_Drawn * 16 + 2) * Screen.TwipsPerPixelY
150       Else
160           Amount_Drawn = itemcount
170           picDropDown.height = (Amount_Drawn * 16 + 16 + 1 + 2) * Screen.TwipsPerPixelY
180       End If

          'change the controls size to fit
190       scroll.height = picDropDown.height - 1 * Screen.TwipsPerPixelY
200       picDropDown.width = UserControl.width - 3 * Screen.TwipsPerPixelX
210       scroll.Left = picDropDown.width - scroll.width

          'show the picdropdown, give it to the desktopwindow, and move it to the right position
          Dim pt As POINTAPI
220       pt.X = -1
230       pt.Y = 17
240       ClientToScreen UserControl.hWnd, pt
250       setParent picDropDown.hWnd, GetDesktopWindow
260       picDropDown.Move ScaleX(pt.X, vbPixels, vbTwips), ScaleY(pt.Y, vbPixels, vbTwips)
270       ShowWindow picDropDown.hWnd, SW_SHOWNOACTIVATE
280       picDropDown.visible = True

290       Redraw
End Sub

Sub HideList()
10        drawFull = False

20        Call UserControl_Resize
30        picDropDown.visible = False

40        Redraw
End Sub



Sub Redraw()
10        UserControl.Cls
20        picDropDown.Cls

          Dim yoffset As Integer

          'draw selected (on usercontrol)
30        If itemcount <> 0 Then
40            piccolor.BackColor = items(ListIndex).color
50            BitBlt UserControl.hDC, 0, yoffset, 16, 16, piccolor.hDC, 0, 0, vbSrcCopy

60            BitBlt UserControl.hDC, 16, yoffset, 16, 16, picicons.hDC, ICON_DELETE, 0, vbSrcCopy
70            BitBlt UserControl.hDC, 32, yoffset, 16, 16, picicons.hDC, ICON_EDIT, 0, vbSrcCopy
80            BitBlt UserControl.hDC, 48, yoffset, 16, 16, picicons.hDC, IIf(items(ListIndex).visible, ICON_VISIBLE, ICON_INVISIBLE), 0, vbSrcCopy


90            UserControl.CurrentX = 65 * Screen.TwipsPerPixelX
100           UserControl.CurrentY = (yoffset + IIf(TextHeight(items(ListIndex).name) / Screen.TwipsPerPixelY < 16, (16 - TextHeight(items(ListIndex).name) / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
110           UserControl.ForeColor = vbBlack
120           UserControl.Print items(ListIndex).name

              'draw a line to seperate from the drop down (not visible if not dropped down)
130           UserControl.Line (0, 16 * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth, 16 * Screen.TwipsPerPixelY), &HC0C0C0

140       End If


150       If itemcount <> 0 And drawFull Then

              Dim Amount_Drawn As Integer
160           If itemcount > Amount_Shown Then
170               Amount_Drawn = Amount_Shown
180           Else
190               Amount_Drawn = itemcount
200           End If

              Dim j As Integer
210           For j = scroll.value To scroll.value + Amount_Drawn - 1
220               If j < itemcount Then

                      'draw icons
230                   BitBlt picDropDown.hDC, 16, yoffset + (j - scroll.value) * 16, 16, 16, picicons.hDC, ICON_DELETE, 0, vbSrcCopy
240                   BitBlt picDropDown.hDC, 32, yoffset + (j - scroll.value) * 16, 16, 16, picicons.hDC, ICON_EDIT, 0, vbSrcCopy
250                   BitBlt picDropDown.hDC, 48, yoffset + (j - scroll.value) * 16, 16, 16, picicons.hDC, IIf(items(j).visible, ICON_VISIBLE, ICON_INVISIBLE), 0, vbSrcCopy

                      'draw the color of the region
260                   piccolor.BackColor = items(j).color
270                   BitBlt picDropDown.hDC, 0, yoffset + (j - scroll.value) * 16, 16, 16, piccolor.hDC, 0, 0, vbSrcCopy

                      'if item is selected, highlight it
280                   If j = ListIndex Then
290                       Call AlphaBlend(picDropDown.hDC, 0, yoffset + (j - scroll.value) * 16, picHighlight.ScaleWidth, picHighlight.ScaleHeight, picHighlight.hDC, 0, 0, picHighlight.ScaleWidth, picHighlight.ScaleHeight, lbc)
300                   End If

                      ' print the name
310                   picDropDown.CurrentX = 65 * Screen.TwipsPerPixelX
320                   picDropDown.CurrentY = (yoffset + (j - scroll.value) * 16 + IIf(TextHeight(items(j).name) / Screen.TwipsPerPixelY < 16, (16 - TextHeight(items(j).name) / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
330                   picDropDown.ForeColor = vbBlack
340                   picDropDown.Print items(j).name
350               End If
360           Next

              ' change the yoffset depending
370           If itemcount > Amount_Shown Then
380               yoffset = (Amount_Drawn - 1) * 16
390           Else
400               yoffset = (Amount_Drawn - 1) * 16 + 16
410           End If
420       End If

430       If itemcount = 0 And Not drawFull Then
              'draw the Add New Region on the usercontrol because there are no regions
440           UserControl.CurrentX = (((UserControl.ScaleWidth - scroll.width) / Screen.TwipsPerPixelX) / 2 - (TextWidth("Add New Region...") / Screen.TwipsPerPixelX) / 2) * Screen.TwipsPerPixelX
450           UserControl.CurrentY = (1 + yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
460           UserControl.ForeColor = vbBlack
470           UserControl.Print "Add New Region..."

480       ElseIf drawFull And scroll.value = scroll.Max Then
490           picDropDown.Line (0, (yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth, (yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY), &HC0C0C0

              'draw the Add New Region
500           picDropDown.CurrentX = (((UserControl.ScaleWidth - scroll.width) / Screen.TwipsPerPixelX) / 2 - (TextWidth("Add New Region...") / Screen.TwipsPerPixelX) / 2) * Screen.TwipsPerPixelX
510           picDropDown.CurrentY = (1 + yoffset + IIf(TextHeight("Add New Region...") / Screen.TwipsPerPixelY < 16, (16 - TextHeight("Add New Region...") / Screen.TwipsPerPixelY) / 2, 0)) * Screen.TwipsPerPixelY
520           picDropDown.ForeColor = vbBlack
530           picDropDown.Print "Add New Region..."
540       End If


End Sub
Sub Clear()
10        Erase items
20        ReDim items(0)
30        itemcount = 0
40        c_listindex = -1
50        cmdDropDown.Enabled = False

60        Redraw
End Sub

Sub Sort()
10        If itemcount = 0 Then Exit Sub

          Dim selitem As item
20        selitem = items(ListIndex)

30        Call Quicksort(items, 0, itemcount - 1)

          'retrieve the index of the current selected item
          Dim i As Integer
40        For i = 0 To itemcount - 1
50            If selitem.name = items(i).name And selitem.color = items(i).color And selitem.visible = items(i).visible Then
60                ListIndex = i
70                Exit For
80            End If
90        Next

100       Redraw
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
10        If Min >= Max Then Exit Sub

          ' Pick a dividing value.
20        i = Int((Max - Min + 1) * Rnd + Min)
30        mid_value = list(i)

          ' Swap the dividing value to the front.
40        list(i) = list(Min)

50        lo = Min
60        hi = Max
70        Do
              ' Look down from hi for a value < mid_value.
80            Do While list(hi).name >= mid_value.name
90                hi = hi - 1
100               If hi <= lo Then Exit Do
110           Loop
120           If hi <= lo Then
130               list(lo) = mid_value
140               Exit Do
150           End If

              ' Swap the lo and hi values.
160           list(lo) = list(hi)

              ' Look up from lo for a value >= mid_value.
170           lo = lo + 1
180           Do While list(lo).name < mid_value.name
190               lo = lo + 1
200               If lo >= hi Then Exit Do
210           Loop
220           If lo >= hi Then
230               lo = hi
240               list(hi) = mid_value
250               Exit Do
260           End If

              ' Swap the lo and hi values.
270           list(hi) = list(lo)
280       Loop

          ' Sort the two sublists.
290       Quicksort list, Min, lo - 1
300       Quicksort list, lo + 1, Max
End Sub


Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
10        ListIndex = c_listindex
End Property

Public Property Let ListIndex(ByVal lstindex As Integer)
    Dim oldidx As Integer
    oldidx = c_listindex
    
10        If itemcount = 0 Then
20            c_listindex = -1
30        Else
40            If lstindex < 0 Or lstindex >= itemcount Then
50                Call Err.Raise(10001, , "The index '" & lstindex & "' is outside the boundaries of the list")
60                Exit Property
70            End If

80            c_listindex = lstindex
90        End If
    If oldidx <> c_listindex Then
        RaiseEvent change
    End If
End Property

Public Property Get ListCount() As Integer
10        ListCount = itemcount
End Property

Public Property Get LinesShown() As Integer
Attribute LinesShown.VB_Description = "Maximum number of lines that are shown when the list is dropped down"
10        LinesShown = Amount_Shown
End Property

Public Property Let LinesShown(ByVal newval As Integer)
10        Amount_Shown = newval
20        PropertyChanged "LinesShown"

End Property

