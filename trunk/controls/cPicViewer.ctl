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
    Animation = m_isAnimation
End Property

Public Property Let Animation(anim As Boolean)
    m_isAnimation = anim
    PropertyChanged "Animation"
    
    CurrentFrame = 0
    If anim Then
        animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)
        animtimer.Enabled = True
    Else
        animtimer.Enabled = False
    End If
    
    Call InitScrollbars
    Call Redraw
End Property

'Time of the animation loop
Public Property Get animationTime() As Long
    animationTime = m_AnimTime
End Property

Public Property Let animationTime(time As Long)
    If time <= 0 Then time = 1
    
    m_AnimTime = time
    
    animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)
    
    PropertyChanged "AnimationTime"
End Property

'Animation frames X/Y
Public Property Get AnimFramesX() As Integer
    AnimFramesX = m_FramesX
End Property

Public Property Let AnimFramesX(FramesX As Integer)
    If FramesX <= 0 Then FramesX = 1
    
    m_FramesX = FramesX

    animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)

    PropertyChanged "AnimFramesX"
    
    
    If m_isAnimation Then Call Redraw
End Property

Public Property Get AnimFramesY() As Integer
    AnimFramesY = m_FramesY
End Property

Public Property Let AnimFramesY(FramesY As Integer)
    If FramesY <= 0 Then FramesY = 1
    
    m_FramesY = FramesY

    animtimer.Interval = m_AnimTime / (m_FramesX * m_FramesY)

    PropertyChanged "AnimFramesY"
    
    If m_isAnimation Then Call Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = picbox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picbox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = picbox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picbox.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = picbox.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    picbox.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Call InitScrollbars
    Call Redraw
End Sub






Private Sub mnuZoomIn_Click()
    Call ZoomIn(CInt(lastclickX), CInt(lastclickY))
End Sub

Private Sub animtimer_Timer()
    CurrentFrame = CurrentFrame + 1
    Call Redraw
End Sub

Private Sub picbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lastclickX = X
    lastclickY = Y
    
    If Button = vbRightButton Then
        
    Else
        'Do stuff
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Private Sub picbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get hDC() As Long
    hDC = picfull.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picfull.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picfull.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = picfull.Image
End Property

Private Sub UserControl_Initialize()
    m_FramesX = 1
    m_FramesY = 1
End Sub

Private Sub UserControl_Resize()
    picbox.Width = UserControl.ScaleWidth - Vscr.Width
    picbox.Height = UserControl.ScaleHeight - Hscr.Height
    
    Vscr.Left = UserControl.ScaleWidth - Vscr.Width
    Vscr.Height = UserControl.ScaleHeight - Hscr.Height
    
    Hscr.Top = UserControl.ScaleHeight - Hscr.Height
    Hscr.Width = UserControl.ScaleWidth - Vscr.Width
    
    Call InitScrollbars
    Call Redraw
    
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picfull,picfull,-1,ScaleWidth
Public Property Get imageWidth() As Single
Attribute imageWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    imageWidth = picfull.Width
End Property

Public Property Let imageWidth(ByVal New_ImageWidth As Single)
    picfull.Width() = New_ImageWidth
    PropertyChanged "ImageWidth"
    
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picbox,picbox,-1,ScaleHeight
Public Property Get imageHeight() As Single
Attribute imageHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    imageHeight = picfull.Height
End Property

Public Property Let imageHeight(ByVal New_ImageHeight As Single)
    picfull.Height() = New_ImageHeight
    PropertyChanged "ImageHeight"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackStyle = m_def_BackStyle
    m_ZoomStep = 1
    m_FramesX = 1
    m_FramesY = 1
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picbox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set picbox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    picbox.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    picfull.Width = PropBag.ReadProperty("ImageWidth", 165)
    picfull.Height = PropBag.ReadProperty("ImageHeight", 117)
    
    m_isAnimation = PropBag.ReadProperty("Animation", False)
    m_AnimTime = PropBag.ReadProperty("AnimationTime", 100)
    m_FramesX = PropBag.ReadProperty("AnimFramesX", 1)
    m_FramesY = PropBag.ReadProperty("AnimFramesY", 1)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", picbox.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", picbox.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", picbox.BorderStyle, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ImageWidth", picfull.Width, 165)
    Call PropBag.WriteProperty("ImageHeight", picfull.Height, 117)
    Call PropBag.WriteProperty("Animation", m_isAnimation, False)
    Call PropBag.WriteProperty("AnimationTime", m_AnimTime, 100)
    Call PropBag.WriteProperty("AnimFramesX", m_FramesX, 1)
    Call PropBag.WriteProperty("AnimFramesY", m_FramesY, 1)
End Sub

Public Sub LoadPicture(path As String)
    picfull.AutoSize = True
    Call LoadPic(picfull, path)
    
    Call InitScrollbars
    Call Redraw
End Sub


Private Sub InitScrollbars()
   
    Dim maxh As Single, maxv As Single
    Dim framew As Integer, frameh As Integer
    
    If m_isAnimation Then
        framew = picfull.Width \ m_FramesX
        frameh = picfull.Height \ m_FramesY
    Else
        framew = picfull.Width
        frameh = picfull.Height
    End If
    
    
    Dim z As Single
    z = zoom()
    maxh = framew - picbox.Width / z
    maxv = frameh - picbox.Height / z
    
    
    If maxh >= 1 Then
        Hscr.Max = RoundAway(maxh)
        Hscr.Enabled = True
    Else
        Hscr.Value = 0
        Hscr.Enabled = False
    End If
    
    If maxv >= 1 Then
        Vscr.Max = RoundAway(maxv)
        Vscr.Enabled = True
    Else
        Vscr.Value = 0
        Vscr.Enabled = False
    End If
  
    
End Sub

Private Sub Redraw()
    Dim z As Single
    picbox.Cls
    
    z = zoom()
    
    Dim SrcX As Integer, SrcY As Integer, srcW As Integer, srcH As Integer

    
    Dim framew As Integer, frameh As Integer
    
  If m_FramesX < 1 Then m_FramesX = 1
  If m_FramesY < 1 Then m_FramesY = 1
      
    If m_isAnimation Then

        framew = picfull.Width \ m_FramesX
        frameh = picfull.Height \ m_FramesY
    Else
        framew = picfull.Width
        frameh = picfull.Height
    End If
    

    
    
    SrcX = (CurrentFrame Mod m_FramesX) * framew
    SrcY = ((CurrentFrame \ m_FramesX) Mod m_FramesY) * frameh
    
    piczoomframe.Cls
    piczoomframe.Width = framew
    piczoomframe.Height = frameh
    
    BitBlt piczoomframe.hDC, 0, 0, piczoomframe.Width, piczoomframe.Height, picfull.hDC, SrcX, SrcY, vbSrcCopy
    
    If Abs(z) = 1# Then
        BitBlt picbox.hDC, 0, 0, picbox.Width, picbox.Height, piczoomframe.hDC, Hscr.Value, Vscr.Value, vbSrcCopy
    Else
        If z > 0 Then
            SetStretchBltMode picbox.hDC, COLORONCOLOR
        Else
            SetStretchBltMode picbox.hDC, HALFTONE
        End If
        
        StretchBlt picbox.hDC, 0, 0, picbox.Width, picbox.Height, piczoomframe.hDC, Hscr.Value, Vscr.Value, picbox.Width / z, picbox.Height / z, vbSrcCopy
    End If
    
    'BitBlt picbox.hDC, 0, 0, picbox.width, picbox.height, piczoomframe.hDC, hscr.Value * z, vscr.Value * z, vbSrcCopy
    
'    srcW = framew / z
'    srcH = frameh / z
'
'    If Abs(z) = 1# Then
'        BitBlt picbox.hDC, 0, 0, framew, frameh, picfull.hDC, SrcX, SrcY, vbSrcCopy
'    Else
'        StretchBlt picbox.hDC, 0, 0, picbox.width, picbox.height, picfull.hDC, SrcX, SrcY, srcW, srcH, vbSrcCopy
'    End If
    

    

    
    picbox.Refresh
End Sub

Private Sub Hscr_Change()
    Call Redraw
End Sub

Private Sub Vscr_Change()
    Call Redraw
End Sub

Private Sub hScr_Scroll()
    Call Redraw
End Sub

Private Sub Vscr_Scroll()
    Call Redraw
End Sub


Public Sub ZoomIn(Optional X As Integer = -1, Optional Y As Integer = -1)
    If m_ZoomStep < MAX_ZOOM Then
        If m_ZoomStep = -1 Then
            m_ZoomStep = 2
        Else
            m_ZoomStep = m_ZoomStep + 1
        End If
    
        Call InitScrollbars
        
        Call Redraw
    End If
    RaiseEvent zoom
    
End Sub

Public Sub ZoomOut(Optional X As Integer = -1, Optional Y As Integer = -1)
    If m_ZoomStep > MIN_ZOOM Then
        If m_ZoomStep = 1 Then
            m_ZoomStep = -2
        Else
            m_ZoomStep = m_ZoomStep - 1
        End If
    
        Call InitScrollbars
        
        Call Redraw
    End If
    
    RaiseEvent zoom
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
    If m_ZoomStep = 0 Then m_ZoomStep = 1
    
    If m_ZoomStep >= 0 Then
        zoom = m_ZoomStep
    Else
        zoom = -1# / m_ZoomStep
    End If
End Property

Public Property Get zoomstr() As String
    If m_ZoomStep = 0 Then m_ZoomStep = 1
    
    If m_ZoomStep >= 0 Then
        zoomstr = m_ZoomStep & ":1"
    Else
        zoomstr = "1:" & Abs(m_ZoomStep)
    End If
End Property

Public Sub Clear()
    picfull.Cls
    picfull.Width = 1
    picfull.Height = 1
    Call InitScrollbars
    Call Redraw
End Sub

Private Function RoundAway(X As Single) As Integer
    RoundAway = Sgn(-X) * Int(-Abs(X))
End Function
