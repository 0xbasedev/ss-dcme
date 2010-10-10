VERSION 5.00
Begin VB.UserControl cPicViewer 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   ToolboxBitmap   =   "cPicViewer.ctx":0000
   Begin VB.PictureBox piczoomframe 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picfull 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.HScrollBar hscr 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2535
   End
   Begin VB.VScrollBar vscr 
      Height          =   1815
      LargeChange     =   10
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picbox 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.Timer animtimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1560
         Top             =   1200
      End
   End
End
Attribute VB_Name = "cPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long


'Default Property Values:
Const m_def_BackStyle = 0
'Property Variables:
Dim m_BackStyle As Integer

'Animation settings
Dim m_isAnimation As Boolean
Dim m_AnimTime As Long
Dim m_FramesX As Integer
Dim m_FramesY As Integer

Dim CurrentFrame As Integer

'm_Zoomstep
'Positive numbers are X:1 ratios; Negative are 1:X ratios
'ex.: if set to 2, it is a 2:1 zoom, but if set to -4, it's a 1:4 zoom
'-1 and 1 are equivalent, 0 is impossible
Dim m_ZoomStep As Integer

'Event Declarations:
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
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event zoom()

Dim lastclickX As Single, lastclickY As Single

Const MAX_ZOOM = 32
Const MIN_ZOOM = -32

'Is an animation?
Public Property Get Animation() As Boolean
10        Animation = m_isAnimation
End Property

Public Property Let Animation(anim As Boolean)
10        m_isAnimation = anim
20        PropertyChanged "Animation"
          
30        CurrentFrame = 0
40        If anim Then
50            animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)
60            animtimer.Enabled = True
70        Else
80            animtimer.Enabled = False
90        End If
          
100       Call InitScrollbars
110       Call Redraw
End Property

'Time of the animation loop
Public Property Get animationTime() As Long
10        animationTime = m_AnimTime
End Property

Public Property Let animationTime(time As Long)
10        If time <= 0 Then time = 1
          
20        m_AnimTime = time
          
30        animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)
          
40        PropertyChanged "AnimationTime"
End Property

'Animation frames X/Y
Public Property Get AnimFramesX() As Integer
10        AnimFramesX = m_FramesX
End Property

Public Property Let AnimFramesX(FramesX As Integer)
10        If FramesX <= 0 Then FramesX = 1
          
20        m_FramesX = FramesX

    animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)

30        PropertyChanged "AnimFramesX"
          
          
40        If m_isAnimation Then Call Redraw
End Property

Public Property Get AnimFramesY() As Integer
10        AnimFramesY = m_FramesY
End Property

Public Property Let AnimFramesY(FramesY As Integer)
10        If FramesY <= 0 Then FramesY = 1
          
20        m_FramesY = FramesY

    animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)

30        PropertyChanged "AnimFramesY"
          
40        If m_isAnimation Then Call Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
10        BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
10        UserControl.BackColor() = New_BackColor
20        PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
10        ForeColor = picbox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
10        picbox.ForeColor() = New_ForeColor
20        PropertyChanged "ForeColor"
End Property

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
10        Set Font = picbox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
10        Set picbox.Font = New_Font
20        PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
10        BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
10        m_BackStyle = New_BackStyle
20        PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
10        BorderStyle = picbox.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
10        picbox.BorderStyle() = New_BorderStyle
20        PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub






Private Sub mnuZoomIn_Click()
10        Call ZoomIn(CInt(lastclickX), CInt(lastclickY))
End Sub

Private Sub animtimer_Timer()
10        CurrentFrame = CurrentFrame + 1
20        Call Redraw
End Sub

Private Sub picbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        lastclickX = X
20        lastclickY = Y
          
30        If Button = vbRightButton Then
              
40        Else
              'Do stuff
50        End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Private Sub picbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
10        RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
10        RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
10        Set Picture = picfull.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
10        Set picfull.Picture = New_Picture
20        PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
10        Set Image = picfull.Image
End Property

Private Sub UserControl_Resize()
10        picbox.width = UserControl.ScaleWidth - vscr.width
20        picbox.height = UserControl.ScaleHeight - hscr.height
          
30        vscr.Left = UserControl.ScaleWidth - vscr.width
40        vscr.height = UserControl.ScaleHeight - hscr.height
          
50        hscr.Top = UserControl.ScaleHeight - hscr.height
60        hscr.width = UserControl.ScaleWidth - vscr.width
          
70        Call InitScrollbars
80        Call Redraw
          
90        RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picfull,picfull,-1,ScaleWidth
Public Property Get imageWidth() As Single
Attribute imageWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
10        imageWidth = picfull.width
End Property

Public Property Let imageWidth(ByVal New_ImageWidth As Single)
10        picfull.width() = New_ImageWidth
20        PropertyChanged "ImageWidth"
          
          
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,ScaleHeight
Public Property Get imageHeight() As Single
Attribute imageHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
10        imageHeight = picbox.height
End Property

Public Property Let imageHeight(ByVal New_ImageHeight As Single)
10        picbox.height() = New_ImageHeight
20        PropertyChanged "ImageHeight"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
10        m_BackStyle = m_def_BackStyle
20        m_ZoomStep = 1
30        m_FramesX = 1
40        m_FramesY = 1
          
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

10        UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
20        picbox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
30        UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
40        Set picbox.Font = PropBag.ReadProperty("Font", Ambient.Font)
50        m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
60        picbox.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
70        Set Picture = PropBag.ReadProperty("Picture", Nothing)
80        picfull.width = PropBag.ReadProperty("ImageWidth", 165)
90        picfull.height = PropBag.ReadProperty("ImageHeight", 117)
          
100       m_isAnimation = PropBag.ReadProperty("Animation", False)
110       m_AnimTime = PropBag.ReadProperty("AnimationTime", 100)
120       m_FramesX = PropBag.ReadProperty("AnimFramesX", 1)
130       m_FramesY = PropBag.ReadProperty("AnimFramesY", 1)
          
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

10        Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
20        Call PropBag.WriteProperty("ForeColor", picbox.ForeColor, &H80000012)
30        Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
40        Call PropBag.WriteProperty("Font", picbox.Font, Ambient.Font)
50        Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
60        Call PropBag.WriteProperty("BorderStyle", picbox.BorderStyle, 0)
70        Call PropBag.WriteProperty("Picture", Picture, Nothing)
80        Call PropBag.WriteProperty("ImageWidth", picfull.width, 165)
90        Call PropBag.WriteProperty("ImageHeight", picfull.height, 117)
100       Call PropBag.WriteProperty("Animation", m_isAnimation, False)
110       Call PropBag.WriteProperty("AnimationTime", m_AnimTime, 100)
120       Call PropBag.WriteProperty("AnimFramesX", m_FramesX, 1)
130       Call PropBag.WriteProperty("AnimFramesY", m_FramesY, 1)
End Sub

Public Sub LoadPicture(path As String)
10        picfull.AutoSize = True
20        Call LoadPic(picfull, path)
          
30        Call InitScrollbars
40        Call Redraw
End Sub


Private Sub InitScrollbars()
         
          Dim maxh As Single, maxv As Single
          Dim framew As Integer, frameh As Integer
          
10        If m_isAnimation Then
20            framew = picfull.width \ m_FramesX
30            frameh = picfull.height \ m_FramesY
40        Else
50            framew = picfull.width
60            frameh = picfull.height
70        End If
          
          
          Dim z As Single
80        z = zoom()
90        maxh = framew - picbox.width / z
100       maxv = frameh - picbox.height / z
          
          
110       If maxh >= 1 Then
120           hscr.Max = RoundAway(maxh)
130           hscr.Enabled = True
140       Else
150           hscr.value = 0
160           hscr.Enabled = False
170       End If
          
180       If maxv >= 1 Then
190           vscr.Max = RoundAway(maxv)
200           vscr.Enabled = True
210       Else
220           vscr.value = 0
230           vscr.Enabled = False
240       End If
        
          
End Sub

Private Sub Redraw()
          Dim z As Single
10        picbox.Cls
          
20        z = zoom()
          
          Dim SrcX As Integer, SrcY As Integer, srcW As Integer, srcH As Integer

          
          Dim framew As Integer, frameh As Integer
          
30        If m_isAnimation Then
40            framew = picfull.width \ m_FramesX
50            frameh = picfull.height \ m_FramesY
60        Else
70            framew = picfull.width
80            frameh = picfull.height
90        End If
          

          
          
100       SrcX = (CurrentFrame Mod m_FramesX) * framew
110       SrcY = ((CurrentFrame \ m_FramesX) Mod m_FramesY) * frameh
          
120       piczoomframe.Cls
130       piczoomframe.width = framew
140       piczoomframe.height = frameh
          
150       BitBlt piczoomframe.hDC, 0, 0, piczoomframe.width, piczoomframe.height, picfull.hDC, SrcX, SrcY, vbSrcCopy
          
160       If Abs(z) = 1# Then
170           BitBlt picbox.hDC, 0, 0, picbox.width, picbox.height, piczoomframe.hDC, hscr.value, vscr.value, vbSrcCopy
180       Else
190           If z > 0 Then
200               SetStretchBltMode picbox.hDC, COLORONCOLOR
210           Else
220               SetStretchBltMode picbox.hDC, HALFTONE
230           End If
              
240           StretchBlt picbox.hDC, 0, 0, picbox.width, picbox.height, piczoomframe.hDC, hscr.value, vscr.value, picbox.width / z, picbox.height / z, vbSrcCopy
250       End If
          
          'BitBlt picbox.hDC, 0, 0, picbox.width, picbox.height, piczoomframe.hDC, hscr.Value * z, vscr.Value * z, vbSrcCopy
          
      '    srcW = framew / z
      '    srcH = frameh / z
      '
      '    If Abs(z) = 1# Then
      '        BitBlt picbox.hDC, 0, 0, framew, frameh, picfull.hDC, SrcX, SrcY, vbSrcCopy
      '    Else
      '        StretchBlt picbox.hDC, 0, 0, picbox.width, picbox.height, picfull.hDC, SrcX, SrcY, srcW, srcH, vbSrcCopy
      '    End If
          

          

          
260       picbox.Refresh
End Sub

Private Sub Hscr_Change()
10        Call Redraw
End Sub

Private Sub Vscr_Change()
10        Call Redraw
End Sub

Private Sub hScr_Scroll()
10        Call Redraw
End Sub

Private Sub Vscr_Scroll()
10        Call Redraw
End Sub


Public Sub ZoomIn(Optional X As Integer = -1, Optional Y As Integer = -1)
10        If m_ZoomStep < MAX_ZOOM Then
20            If m_ZoomStep = -1 Then
30                m_ZoomStep = 2
40            Else
50                m_ZoomStep = m_ZoomStep + 1
60            End If
          
70            Call InitScrollbars
              
80            Call Redraw
90        End If
100       RaiseEvent zoom
          
End Sub

Public Sub ZoomOut(Optional X As Integer = -1, Optional Y As Integer = -1)
10        If m_ZoomStep > MIN_ZOOM Then
20            If m_ZoomStep = 1 Then
30                m_ZoomStep = -2
40            Else
50                m_ZoomStep = m_ZoomStep - 1
60            End If
          
70            Call InitScrollbars
              
80            Call Redraw
90        End If
          
100       RaiseEvent zoom
End Sub


'    If Not out Then
'        Call Maps(activemap).magnifier.ZoomIn( _
'             (Int(((Maps(activemap).hscr.Value + Maps(activemap).picPreview.width / 2)) / (TileW * Maps(activemap).magnifier.Zoom))), _
'             (Int(((Maps(activemap).vscr.Value + Maps(activemap).picPreview.height / 2)) / (TileW * Maps(activemap).magnifier.Zoom))))
'    Else
'        Call Maps(activemap).magnifier.ZoomOut( _
'             (Int(((Maps(activemap).hscr.Value + Maps(activemap).picPreview.width / 2)) / (TileW * Maps(activemap).magnifier.Zoom))), _
'             (Int(((Maps(activemap).vscr.Value + Maps(activemap).picPreview.height / 2)) / (TileW * Maps(activemap).magnifier.Zoom))))
'    End If


Public Property Get zoom() As Single
10        If m_ZoomStep = 0 Then m_ZoomStep = 1
          
20        If m_ZoomStep >= 0 Then
30            zoom = m_ZoomStep
40        Else
50            zoom = -1# / m_ZoomStep
60        End If
End Property

Public Property Get zoomstr() As String
10        If m_ZoomStep = 0 Then m_ZoomStep = 1
          
20        If m_ZoomStep >= 0 Then
30            zoomstr = m_ZoomStep & ":1"
40        Else
50            zoomstr = "1:" & Abs(m_ZoomStep)
60        End If
End Property

Public Sub Clear()
10        picfull.Cls
20        picfull.width = 1
30        picfull.height = 1
40        Call InitScrollbars
50        Call Redraw
End Sub

Private Function RoundAway(X As Single) As Integer
10        RoundAway = Sgn(-X) * Int(-Abs(X))
End Function
