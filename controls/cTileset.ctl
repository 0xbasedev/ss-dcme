VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl cTileset 
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picTilesetView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   1
      Left            =   7200
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   8
      Top             =   3720
      Width           =   4575
      Begin VB.PictureBox picextratiles 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   0
         Picture         =   "cTileset.ctx":0000
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   304
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.PictureBox pictileset 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3360
         Left            =   0
         ScaleHeight     =   224
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   304
         TabIndex        =   9
         Top             =   0
         Width           =   4560
         Begin VB.Shape shptilesel 
            BorderColor     =   &H0000FFFF&
            Height          =   240
            Index           =   2
            Left            =   240
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Shape shptilesel 
            BorderColor     =   &H000000FF&
            Height          =   240
            Index           =   1
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.PictureBox picTilesetView 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   3
      Left            =   1680
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   4650
      Begin VB.PictureBox picLVZImages 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2880
         Left            =   0
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   304
         TabIndex        =   6
         Top             =   0
         Width           =   4560
         Begin VB.VScrollBar vscrLVZImages 
            Height          =   2880
            LargeChange     =   4
            Left            =   4320
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VB.Timer tmrAnimateLVZImagesPreview 
            Interval        =   25
            Left            =   840
            Top             =   2400
         End
         Begin VB.Shape shpLVZsel 
            BorderColor     =   &H000000FF&
            Height          =   480
            Index           =   1
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Shape shpLVZsel 
            BorderColor     =   &H0000FFFF&
            Height          =   480
            Index           =   2
            Left            =   480
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.ComboBox cmbLVZTilesetDisplayType 
         Height          =   315
         ItemData        =   "cTileset.ctx":E442
         Left            =   2760
         List            =   "cTileset.ctx":E458
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3000
         Width           =   1695
      End
      Begin VB.ComboBox cmbLVZTilesetLayerType 
         Height          =   315
         ItemData        =   "cTileset.ctx":E4A2
         Left            =   960
         List            =   "cTileset.ctx":E4BE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
      End
   End
   Begin VB.PictureBox picTilesetView 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   2
      Left            =   7560
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   4650
      Begin VB.PictureBox picWalltiles 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   0
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   2
         Top             =   0
         Width           =   3840
         Begin VB.Shape shpwallsel 
            BorderColor     =   &H000000FF&
            Height          =   960
            Index           =   1
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Shape shpwallsel 
            BorderColor     =   &H0000FFFF&
            Height          =   960
            Index           =   2
            Left            =   1680
            Top             =   360
            Visible         =   0   'False
            Width           =   960
         End
      End
   End
   Begin MSComctlLib.TabStrip tbTilesetView 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tileset"
            Key             =   "Tileset"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Walltiles"
            Key             =   "Walltiles"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LVZ"
            Key             =   "LVZ"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Shape tmpShape 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   2160
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "cTileset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ToolTipText = ""

Const m_def_LeftSelectionColor = DEFAULT_LEFTCOLOR
Const m_def_RightSelectionColor = DEFAULT_RIGHTCOLOR
Const m_def_EnableLeft = True
Const m_def_EnableRight = True
Const m_def_ShowTilesetTab = True
Const m_def_ShowWalltilesTab = True
Const m_def_ShowLvzTab = True

'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_ToolTipText As String

Dim m_LeftSelectionColor As Long
Dim m_RightSelectionColor As Long
Dim m_EnableLeft As Boolean
Dim m_EnableRight As Boolean
Dim m_ShowTilesetTab As Boolean
Dim m_ShowWalltilesTab As Boolean
Dim m_ShowLvzTab As Boolean



'Event Declarations:
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
Event SelectionChange()
Event TabChange()
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."


Private Type ShapeData
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
    color As Long
    
    OnTab As TilesetTabs 'Tab on which the shape should be shown
End Type












Dim shapes(1 To 2) As ShapeData


Dim CurrentTab As TilesetTabs


Dim parent As frmMain

Dim WithEvents CurrentSelections As TilesetSelections
Attribute CurrentSelections.VB_VarHelpID = -1



'holds the x and y of the tile from the tileset
Dim oldtilesetX As Integer
Dim oldtilesetY As Integer









'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
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
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get tooltiptext() As String
Attribute tooltiptext.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    tooltiptext = m_ToolTipText
End Property

Public Property Let tooltiptext(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property


Public Property Get LeftColor() As Long
    LeftColor = m_LeftSelectionColor
End Property

Public Property Let LeftColor(ByVal New_LeftColor As Long)
    m_LeftSelectionColor = New_LeftColor
    PropertyChanged "LeftColor"
End Property

Public Property Get RightColor() As Long
    RightColor = m_RightSelectionColor
End Property

Public Property Let RightColor(ByVal New_RightColor As Long)
    m_RightSelectionColor = New_RightColor
    PropertyChanged "RightColor"
End Property





Private Sub CurrentSelections_OnChange()
    Call Redraw
    
    RaiseEvent SelectionChange
End Sub





'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ToolTipText = m_def_ToolTipText
    
    
    m_LeftSelectionColor = m_def_LeftSelectionColor
    m_RightSelectionColor = m_def_RightSelectionColor
    m_EnableLeft = m_def_EnableLeft
    m_EnableRight = m_def_EnableRight
    m_ShowTilesetTab = m_def_ShowTilesetTab
    m_ShowWalltilesTab = m_def_ShowWalltilesTab
    m_ShowLvzTab = m_def_ShowLvzTab
    

    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    
    m_LeftSelectionColor = PropBag.ReadProperty("LeftSelectionColor", m_def_LeftSelectionColor)
    m_RightSelectionColor = PropBag.ReadProperty("RightSelectionColor", m_def_RightSelectionColor)
    m_EnableLeft = PropBag.ReadProperty("EnableLeft", m_def_EnableLeft)
    m_EnableRight = PropBag.ReadProperty("EnableRight", m_def_EnableRight)
    m_ShowTilesetTab = PropBag.ReadProperty("ShowTilesetTab", m_def_ShowTilesetTab)
    m_ShowWalltilesTab = PropBag.ReadProperty("ShowWalltilesTab", m_def_ShowWalltilesTab)
    m_ShowLvzTab = PropBag.ReadProperty("ShowLvzTab", m_def_ShowLvzTab)
    
    m_LeftSelectionColor = PropBag.ReadProperty("LeftColor", vbRed)
    m_RightSelectionColor = PropBag.ReadProperty("RightColor", vbYellow)
    
End Sub



'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    
    
    Call PropBag.WriteProperty("LeftSelectionColor", m_LeftSelectionColor, m_def_LeftSelectionColor)
    Call PropBag.WriteProperty("RightSelectionColor", m_RightSelectionColor, m_def_RightSelectionColor)
    Call PropBag.WriteProperty("EnableLeft", m_EnableLeft, m_def_EnableLeft)
    Call PropBag.WriteProperty("EnableRight", m_EnableRight, m_def_EnableRight)
    Call PropBag.WriteProperty("ShowTilesetTab", m_ShowTilesetTab, m_def_ShowTilesetTab)
    Call PropBag.WriteProperty("ShowWalltilesTab", m_ShowWalltilesTab, m_def_ShowWalltilesTab)
    Call PropBag.WriteProperty("ShowLvzTab", m_ShowLvzTab, m_def_ShowLvzTab)
    
    Call PropBag.WriteProperty("LeftColor", m_LeftSelectionColor, vbRed)
    Call PropBag.WriteProperty("RightColor", m_RightSelectionColor, vbYellow)
    
    
End Sub


Private Sub UserControl_Terminate()
    Set CurrentSelections = Nothing
    Set parent = Nothing
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Redraw()
     
    Dim i As Integer
    
    If Not parent Is Nothing Then
        CurrentTab = parent.tileset.CurrentTab
    
        tbTilesetView.Tabs(CurrentTab).selected = True
    End If
    
    If Not CurrentSelections Is Nothing Then
    
        'Update the shapes data
        
        For i = 1 To 2
            
            Select Case CurrentSelections.selection(i).selectionType
            Case TS_Tiles
                Call RefreshTileSelection(i)
            Case TS_Walltiles
                Call RefreshWalltileSelection(i)
            Case TS_LVZ
                Call RefreshLvzSelection(i)
                
            Case Else
                'Do nothing
                
            End Select
            
        Next
        
        
        'Make the shapes white if needed, else reset their colors
        Call CheckShapeColors(shapes(1), shapes(2))
    
        'Apply the shapes data to the actual selection rectangles
        For i = 1 To 2
            
            
            Select Case CurrentTab
            Case TB_Tiles
                Call ApplyShapedataToRectangle(shapes(i), shptilesel(i))
            Case TB_Walltiles
                Call ApplyShapedataToRectangle(shapes(i), shpwallsel(i))
            Case TB_LVZ
                Call ApplyShapedataToRectangle(shapes(i), shpLVZsel(i))
                
            Case Else
                'Do nothing
                
            End Select

            
            shptilesel(i).visible = (shapes(i).OnTab = TB_Tiles)
            shpwallsel(i).visible = (shapes(i).OnTab = TB_Walltiles)
            shpLVZsel(i).visible = (shapes(i).OnTab = TB_LVZ)
            
        Next
    End If
     
End Sub

Private Sub ApplyShapedataToRectangle(ByRef shpdata As ShapeData, ByRef shp As shape)

    shp.Left = shpdata.Left
    shp.Top = shpdata.Top
    shp.width = shpdata.width
    shp.height = shpdata.height
    
    shp.BorderColor = shpdata.color
End Sub


Public Sub setParent(ByRef Main As frmMain)
    Set parent = Main
    
    If Not parent Is Nothing Then
        Set CurrentSelections = Main.tileset
        
        Call DrawWalltilesTileset
    Else
        Set CurrentSelections = Nothing
    
        Call ClearTileset
    End If
    
    Call Redraw
End Sub

Public Sub ClearTileset()
    pictileset.Cls
    picWalltiles.Cls
    picLVZImages.Cls
End Sub


Private Sub RefreshTileSelection(Button As Integer)
    
    With shapes(Button)
    
        .Left = ((CurrentSelections.selection(Button).tilenr - 1) Mod 19) * TILEW
        .Top = ((CurrentSelections.selection(Button).tilenr - 1) \ 19) * TILEH
        
        If CurrentSelections.selection(Button).isSpecialObject Then
            .width = TILEW
            .height = TILEH
        Else
            .width = CurrentSelections.selection(Button).pixelSize.X
            .height = CurrentSelections.selection(Button).pixelSize.Y
        End If
    
        .OnTab = TB_Tiles
    End With
End Sub



Private Sub RefreshWalltileSelection(Button As Integer)
    'Place the selection rectangle
    With shapes(Button)
        .Left = WT_SETW * (CurrentSelections.selection(Button).group Mod WT_NTILESW)
        .Top = WT_SETH * (CurrentSelections.selection(Button).group \ WT_NTILESW)
        
        .width = WT_SETW
        .height = WT_SETH
        
        .OnTab = TB_Walltiles
    End With

End Sub


Private Sub RefreshLvzSelection(Button As Integer)
    With shapes(Button)
        'Show this one

        .Left = (CurrentSelections.LvzLastClickX \ ImageTilesetW) * ImageTilesetW
        .Top = (CurrentSelections.LvzLastClickY \ ImageTilesetH) * ImageTilesetH
        
        .width = ImageTilesetW
        .height = ImageTilesetH
        
        .OnTab = TB_LVZ

    End With
End Sub


Private Sub CheckShapeColors(ByRef Shape1 As ShapeData, ByRef shape2 As ShapeData)
    If ShapesOverlap(Shape1, shape2) Then
        Shape1.color = vbWhite
        shape2.color = vbWhite
    Else
        'Default colors
        Shape1.color = LeftColor
        shape2.color = RightColor
    End If
End Sub




Private Function ShapesOverlap(ByRef Shape1 As ShapeData, ByRef shape2 As ShapeData) As Boolean
    ShapesOverlap = False
    If Shape1.OnTab = shape2.OnTab Then
        If Shape1.Left = shape2.Left And Shape1.Top = shape2.Top Then
            If Shape1.width = shape2.width And Shape1.height = shape2.height Then
                ShapesOverlap = True
            End If
        End If
    End If
End Function




Public Property Get LeftSelectionColor() As Long

    LeftSelectionColor = m_LeftSelectionColor

End Property

Public Property Let LeftSelectionColor(ByVal lm_LeftSelectionColor As Long)

    m_LeftSelectionColor = lm_LeftSelectionColor

    Call UserControl.PropertyChanged("LeftSelectionColor")

End Property

Public Property Get RightSelectionColor() As Long

    RightSelectionColor = m_RightSelectionColor

End Property

Public Property Let RightSelectionColor(ByVal lm_RightSelectionColor As Long)

    m_RightSelectionColor = lm_RightSelectionColor

    Call UserControl.PropertyChanged("RightSelectionColor")

End Property

Public Property Get EnableLeft() As Boolean

    EnableLeft = m_EnableLeft

End Property

Public Property Let EnableLeft(ByVal bm_EnableLeft As Boolean)

    m_EnableLeft = bm_EnableLeft

    Call UserControl.PropertyChanged("EnableLeft")

End Property

Public Property Get EnableRight() As Boolean

    EnableRight = m_EnableRight

End Property

Public Property Let EnableRight(ByVal bm_EnableRight As Boolean)

    m_EnableRight = bm_EnableRight

    Call UserControl.PropertyChanged("EnableRight")

End Property

Public Property Get ShowTilesetTab() As Boolean

    ShowTilesetTab = m_ShowTilesetTab

End Property

Public Property Let ShowTilesetTab(ByVal bm_ShowTilesetTab As Boolean)

    m_ShowTilesetTab = bm_ShowTilesetTab

    Call UserControl.PropertyChanged("ShowTilesetTab")

End Property

Public Property Get ShowWalltilesTab() As Boolean

    ShowWalltilesTab = m_ShowWalltilesTab

End Property

Public Property Let ShowWalltilesTab(ByVal bm_ShowWalltilesTab As Boolean)

    m_ShowWalltilesTab = bm_ShowWalltilesTab

    Call UserControl.PropertyChanged("ShowWalltilesTab")

End Property

Public Property Get ShowLvzTab() As Boolean

    ShowLvzTab = m_ShowLvzTab

End Property

Public Property Let ShowLvzTab(ByVal bm_ShowLvzTab As Boolean)

    m_ShowLvzTab = bm_ShowLvzTab

    Call UserControl.PropertyChanged("ShowLvzTab")

End Property




'''''''''''''''''''''''''''''''''''''
'LVZ default settings drop down lists

Private Sub cmbLVZTilesetDisplayType_Click()
    If Not parent Is Nothing Then
        parent.lvz.MapObjectDefaultMode = cmbLVZTilesetDisplayType.ListIndex
    End If
End Sub

Private Sub cmbLVZTilesetLayerType_Click()
    If Not parent Is Nothing Then
        parent.lvz.MapObjectDefaultLayer = cmbLVZTilesetLayerType.ListIndex
    End If
End Sub


''''''''''''''''''''''''''''
'LVZ picturebox events
Private Sub picLVZImages_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not parent Is Nothing Then
           
        If Button = vbRightButton Then
            'Select the lvz image before popping up the menu
            Call picLVZImages_MouseMove(vbLeftButton, 0, X, Y)
                       
            PopupMenu frmGeneral.mnuLvz
        Else
            Call picLVZImages_MouseMove(Button, Shift, X, Y)
        End If
        
    End If
End Sub

Private Sub picLVZImages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not parent Is Nothing Then

        Dim lvzidx As Integer, imgidx As Integer
        
        
        If parent.lvz.getImageDefFromCoordinates(CInt(X), CInt(Y), lvzidx, imgidx, (picLVZImages.ScaleWidth - vscrLVZImages.width) + 1, picLVZImages.ScaleHeight, vscrLVZImages.value) <> -1 Then
        
            picLVZImages.tooltiptext = parent.lvz.getLVZ(lvzidx).name & " -> " & _
                                       parent.lvz.getImageDefinition(lvzidx, imgidx).imagename & _
                                       " (" & parent.lvz.getImageWidth(lvzidx, imgidx) & Chr(215) & _
                                       parent.lvz.getImageHeight(lvzidx, imgidx) & ")"
                
    
            If Button > 0 Then
                Call parent.tileset.SelectLVZ(Button, lvzidx, imgidx, True, CInt(X), CInt(Y))
                
            End If
            
            frmGeneral.mnuLvzDeleteImage.Enabled = True
            frmGeneral.mnuLvzEditImage.Enabled = True
            frmGeneral.mnuLvzEditAnimation.Enabled = True
            

                
        Else
            picLVZImages.tooltiptext = ""
            frmGeneral.mnuLvzDeleteImage.Enabled = False
            frmGeneral.mnuLvzEditImage.Enabled = False
            frmGeneral.mnuLvzEditAnimation.Enabled = False
        End If
        
    End If
End Sub


Private Sub vscrLVZImages_Change()
    If Not parent Is Nothing Then
        Call DrawLVZTileset(True)
'        Call parent.tileset.DrawLVZTileset(True)
    End If
End Sub

Private Sub vscrLVZImages_Scroll()
    If Not parent Is Nothing Then
        Call DrawLVZTileset(True)
'        Call parent.tileset.DrawLVZTileset(True)
    End If
End Sub









Sub DrawLVZTileset(drawAll As Boolean)
'    Dim lvz() As LVZstruct
'
'    lvz = parent.lvz.getLVZData
    
    If parent Is Nothing Then
        picLVZImages.Cls
        Exit Sub
    End If
    

    Dim X As Integer, Y As Integer
    Dim frameX As Integer, frameY As Integer
    
    Dim Button As Integer
    
'    frmGeneral.picLVZImages.Cls
    
'    SetStretchBltMode frmGeneral.picLVZImages.hDC, COLORONCOLOR
    
'    Dim SelLVZIdx(1 To 2) As Integer
'    Dim SelImgIdx(1 To 2) As Integer
'    Dim lSelLVZIdx As Integer
'    Dim rSelLVZIdx As Integer
'    Dim lSelImgIdx As Integer
'    Dim rSelImgIdx As Integer
    
'    For button = vbLeftButton To vbRightButton
'        With selections(button)
'            If .selectionType = TS_LVZ Then
'                SelLVZIdx(button) = .group
'                SelImgIdx(button) = .tilenr
'            End If
'        End With
'    Next
    
'    lSelLVZIdx = selImageLeft \ 65536
'    rSelLVZIdx = selImageRight \ 65536
'    lSelImgIdx = selImageLeft Mod 65536
'    rSelImgIdx = selImageRight Mod 65536
    
    Dim totalWidth As Integer
    totalWidth = (picLVZImages.ScaleWidth - vscrLVZImages.width) + 1
    Dim totalHeight As Integer
    totalHeight = (picLVZImages.ScaleHeight - ImageTilesetH)
    
    Dim frameWidth As Integer
    Dim frameHeight As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    Dim count As Integer
    
    Dim maxFrames As Integer
    
    
    Dim frameTime As Long, frameChange As Long
    
    Dim currentTick As Long
    currentTick = GetTickCount
    
    
    If drawAll Then picLVZImages.Cls
    
    
    X = 0
    Y = -vscrLVZImages.value * ImageTilesetW
    
    Dim lvzcount As Integer
    lvzcount = parent.lvz.getLVZCount
    
    Dim imgdef As LVZImageDefinition
    
    For i = 0 To lvzcount - 1
        For j = 0 To parent.lvz.getImageDefinitionCount(i) - 1 'lvz(i).ImageDefinitionCount - 1
            
            imgdef = parent.lvz.getImageDefinition(i, j)
            
            With imgdef
            
'                For button = vbLeftButton To vbRightButton
'                    If i = SelLVZIdx(button) And j = SelImgIdx(button) Then
'                        frmGeneral.shpLVZsel(button).Left = X
'                        frmGeneral.shpLVZsel(button).Top = Y
'                    End If
'                Next
    
                maxFrames = intMinimum(.animationFramesX * .animationFramesY, _
                                       16384)
    '
    '            If lvz(i).imagedefinitions(j).animationFramesX * lvz(i).imagedefinitions(j).animationFramesY - 1 < 1000 Then
    '                maxFrames = lvz(i).imagedefinitions(j).animationFramesX * lvz(i).imagedefinitions(j).animationFramesY - 1
    '            Else
    '                maxFrames = 1000
    '            End If
    
                If maxFrames <= 1 Then
                    frameX = 0
                    frameY = 0
                Else
                    frameX = ((.CurrentFrame Mod maxFrames) Mod .animationFramesX)
                    frameY = (((.CurrentFrame Mod maxFrames) \ .animationFramesX) Mod .animationFramesY)
                    
                    frameTime = (.animationTime * 10#) \ maxFrames
                    
                    If frameTime > 0 Then
                        frameChange = (currentTick - .lastFrameChange) \ frameTime
                    
                        If frameChange > 0 Then
                            Call parent.lvz.UpdateImageAnimation(i, j, _
                                                            .lastFrameChange + frameChange * frameTime, _
                                                            (.CurrentFrame + frameChange) Mod maxFrames)

                        End If
                    End If
                End If
            
                    
                If frameChange Or drawAll Then BitBlt picLVZImages.hDC, X, Y, ImageTilesetW, ImageTilesetH, parent.lvz.pichDClib(.picboxIdx), frameX * ImageTilesetW, frameY * ImageTilesetH, vbSrcCopy
                
'                frameWidth = .picWidth     ' parent.lvz.getImageWidth(i, j)
'                frameHeight = .picHeight   ' parent.lvz.getImageHeight(i, j)
'
'
'
'                If frameWidth <= ImageTilesetW And frameHeight <= ImageTilesetH Then
'                    'Image smaller than the frame, bitblt 1:1
'                    BitBlt frmGeneral.picLVZImages.hDC, (ImageTilesetW - frameWidth) \ 2 + X, (ImageTilesetH - frameHeight) \ 2 + Y, frameWidth, frameHeight, parent.picLVZItem(.picboxIdx).hDC, frameX * frameWidth, frameY * frameHeight, vbSrcCopy
'
'                Else
'
'                    wFactor = frameWidth / ImageTilesetW
'                    hFactor = frameHeight / ImageTilesetH
'
'                    If hFactor > wFactor Then
'                        previewWidth = intMaximum(frameWidth / hFactor, 1)
'                        previewHeight = ImageTilesetH
'
'
'                        SetStretchBltMode frmGeneral.picLVZImages.hDC, IIf(hFactor >= 3, COLORONCOLOR, HALFTONE)
'                    Else
'                        previewWidth = ImageTilesetW
'                        previewHeight = intMaximum(frameHeight / wFactor, 1)
'
'                        SetStretchBltMode frmGeneral.picLVZImages.hDC, IIf(wFactor >= 3, COLORONCOLOR, HALFTONE)
'                    End If
'
'                    StretchBlt frmGeneral.picLVZImages.hDC, (ImageTilesetW - previewWidth) \ 2 + X, (ImageTilesetH - previewHeight) \ 2 + Y, previewWidth, previewHeight, parent.picLVZItem(.picboxIdx).hDC, frameX * frameWidth, frameY * frameHeight, frameWidth, frameHeight, vbSrcCopy
'                End If
  
            End With
      
            count = count + 1
            X = (X + ImageTilesetW) Mod totalWidth
            Y = (count \ (totalWidth \ ImageTilesetW)) * ImageTilesetH - vscrLVZImages.value * ImageTilesetW
            
        Next
    Next
    
    'Samapico: This caused flickering while trying to change layer or display mode
    'frmGeneral.cmbLVZTilesetDisplayType.ListIndex = selDisplay
    'frmGeneral.cmbLVZTilesetLayerType.ListIndex = selLayer
    
    'total amount of images
    count = (count - 1)
    
    Dim maxvscroll As Long
    maxvscroll = (count * ImageTilesetW) \ totalWidth - ((totalHeight - ImageTilesetH) \ ImageTilesetH)
    
    If maxvscroll > 0 Then
        vscrLVZImages.Max = maxvscroll
        vscrLVZImages.Enabled = True
    Else
        vscrLVZImages.Enabled = False
    End If
    
    picLVZImages.Refresh

End Sub



Public Property Let DefaultLvzLayer(new_defaultlayer As LVZLayerEnum)
    cmbLVZTilesetLayerType.ListIndex = new_defaultlayer
End Property

Public Property Let DefaultLvzMode(new_mode As LVZModeEnum)
    cmbLVZTilesetDisplayType.ListIndex = new_mode
End Property

Public Property Get DefaultLvzLayer() As LVZLayerEnum
    DefaultLvzLayer = cmbLVZTilesetLayerType.ListIndex
End Property

Public Property Get DefaultLvzMode() As LVZModeEnum
    DefaultLvzMode = cmbLVZTilesetDisplayType.ListIndex
End Property



Public Property Get Pic_ExtraTiles() As PictureBox
    Set Pic_ExtraTiles = picextratiles
End Property

Public Property Get Pic_WallTiles() As PictureBox
    Set Pic_WallTiles = picWalltiles
End Property

Public Property Get Pic_LvzTiles() As PictureBox
    Set Pic_LvzTiles = picLVZImages
End Property

Public Property Get Pic_Tileset() As PictureBox
    Set Pic_Tileset = pictileset
End Property



Private Sub tmrAnimateLVZImagesPreview_Timer()
    If Not parent Is Nothing Then
        
        If CurrentTab = TB_LVZ Then
            Call DrawLVZTileset(False)
        End If
    'Call Maps(activemap).lvz.DrawLVZImageInterface(Maps(activemap).lvz.curImageFrame)
    'Maps(activemap).lvz.curImageFrame = (Maps(activemap).lvz.curImageFrame + 1) Mod 1000
    End If
End Sub






'''''''''''''''''''''''''''''
'Pictileset events


Sub pictileset_KeyDown(KeyCode As Integer, Shift As Integer)
'Catch hotkeys
    If Not parent Is Nothing Then

        Call parent.UseShortcutTool(True, KeyCode, Shift)

    End If

End Sub

Sub pictileset_KeyUp(KeyCode As Integer, Shift As Integer)
'Catch hotkeys
    If Not parent Is Nothing Then

        Call parent.UseShortcutTool(False, KeyCode, Shift)

    End If
End Sub

Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not parent Is Nothing Then

        'Selects a tile from the tileset
        If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
            'not in the boundaries of the picture
        
        Else
    
            'indicate the mouse is down
            SharedVar.MouseDown = Button
        
        
        
            oldtilesetX = Int((X / 16) + 1)
            oldtilesetY = Int(Y / 16)
        
            'set the selected tile
            If Button = vbLeftButton Or Button = vbRightButton Then
                Call parent.tileset.SelectTiles(Button, oldtilesetY * 19 + oldtilesetX, 1, 1, True)
                
                RaiseEvent SelectionChange
            End If

        End If
    End If
End Sub

Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim tilenr As Integer, sizeX As Integer, sizeY As Integer, curtilesetX As Integer, curtilesetY As Integer
    
    
    
    If Not parent Is Nothing Then

        'Still select a tile from the tileset if mousedown, but show the
        'tooltiptext of the different tiles
        If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
            'not in the boundaries of the picture
        
        
        Else

            tilenr = Int(Y / 16) * 19 + Int((X / 16) + 1)
        
        
            ' needed to move the tooltip if it hadn't changed
            'pictileset.ToolTipText = ""
            
            pictileset.tooltiptext = TilesetToolTipText(tilenr)
        
        
            If Button = vbLeftButton Or Button = vbRightButton Then
        
                
                
                curtilesetX = Int((X / 16) + 1)
                curtilesetY = Y \ 16
        
                sizeX = Abs(curtilesetX - oldtilesetX)
                sizeY = Abs(curtilesetY - oldtilesetY)
        
                If curtilesetX > oldtilesetX Then
                    curtilesetX = oldtilesetX
                End If
                If curtilesetY > oldtilesetY Then
                    curtilesetY = oldtilesetY
                End If
        
                If (curtilesetX <= 8 And 8 <= curtilesetX + sizeX) And _
                   (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
                    sizeX = 0
                    sizeY = 0
                End If
                If (curtilesetX <= 10 And 10 <= curtilesetX + sizeX) And _
                   (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
                    sizeX = 0
                    sizeY = 0
                End If
                If (9 <= curtilesetX + sizeX) And _
                   (curtilesetY <= 13 And 13 <= curtilesetY + sizeY) Then
                    sizeX = 0
                    sizeY = 0
                End If
                If (curtilesetX <= 11 And 11 <= curtilesetX + sizeX) And _
                   (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
                    sizeX = 0
                    sizeY = 0
                End If
                If tilenr > 256 Then
                    sizeX = 0
                    sizeY = 0
                End If
        
        
                If sizeX = 0 Then curtilesetX = oldtilesetX
                If sizeY = 0 Then curtilesetY = oldtilesetY
        
                'set the selected tile
                Call parent.tileset.SelectTiles(Button, curtilesetY * 19 + curtilesetX, sizeX + 1, sizeY + 1, True)
                
                RaiseEvent SelectionChange
            End If
    
        End If
    End If
End Sub

Sub pictileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'mouse is no longer down
    SharedVar.MouseDown = 0

    'try to set focus to the activemap again
'    On Error Resume Next
'    If loadedmaps(activemap) = True And AutoHideTileset = False Then
'        Maps(activemap).picPreview.setfocus
'    End If
End Sub







''''''''''''''''''''''''''''
'picwalltiles events

Private Sub picWalltiles_DblClick()
    If Not parent Is Nothing Then Call frmGeneral.DoEditWalltiles
End Sub

Private Sub picWalltiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not parent Is Nothing Then
        Dim newset As Integer
        newset = (X \ WT_SETW) + (Y \ WT_SETH) * 4
        
        If newset < 0 Or newset > 7 Then Exit Sub
        
        If parent.walltiles.isValidSet(newset) Then
            Call parent.tileset.SelectWalltiles(Button, newset, True)
            
            RaiseEvent SelectionChange
        End If
    End If
    
End Sub

Private Sub picWalltiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not parent Is Nothing Then
        picWalltiles.tooltiptext = "Walltiles set #" & (X \ 64) + (Y \ 64) * 4
    
        If Button Then Call picWalltiles_MouseDown(Button, Shift, X, Y)
    Else
        picWalltiles.tooltiptext = ""
    End If
End Sub





Sub DrawWalltilesTileset()
    If Not parent Is Nothing Then
        Call parent.walltiles.DrawWallTiles(picWalltiles.hDC, 4)
    Else
        picWalltiles.Cls
    End If
    
    picWalltiles.Refresh
End Sub


'''''''''''''''''''''''''''''''''''''
'Tabs

Sub tbTilesetView_Click()
    
    
    Dim i As TilesetTabs
    
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = TilesetTabs.TB_Tiles To TilesetTabs.TB_LVZ
        With picTilesetView(i)
            If i = tbTilesetView.SelectedItem.Index Then
                
                
                .Left = tbTilesetView.Left + 1
                .Enabled = True
                .Top = tbTilesetView.Top + 23
                .visible = True
                .ZOrder vbBringToFront
                
                If Not parent Is Nothing Then
                    With parent
                    
                        .tileset.CurrentTab = i
                        CurrentTab = i
                            
                            
                        If i = TilesetTabs.TB_LVZ Then

                            Call DrawLVZTileset(True)
                            
                            tmrAnimateLVZImagesPreview.Enabled = CBool(GetSetting("AnimatedLVZImageTiles", 1))
            
                            cmbLVZTilesetLayerType.ListIndex = .lvz.MapObjectDefaultLayer
                            cmbLVZTilesetDisplayType.ListIndex = .lvz.MapObjectDefaultMode
                        Else
                            tmrAnimateLVZImagesPreview.Enabled = False
                        End If
                    End With
                End If
                
            Else
                .visible = False
                .Left = -5000
                .Enabled = False
            End If
        End With
    Next
    

    RaiseEvent TabChange
    
End Sub



Sub CopyTilesetFromDC(srcDC As Long, SrcX As Integer, SrcY As Integer)
    BitBlt pictileset.hDC, 0, 0, pictileset.width, pictileset.height, srcDC, SrcX, SrcY, vbSrcCopy
    pictileset.Refresh
End Sub




Sub DrawTilePreview(Button As Integer, ByRef destpic As PictureBox, DestX As Long, DestY As Long, destw As Long, desth As Long)
    'Updates the preview of the tileset
    'draw the large preview of the selected tiles
    
    If parent Is Nothing Then Exit Sub
    
    With parent
        
        tmpShape.Left = DestX
        tmpShape.Top = DestY
        tmpShape.width = destw
        tmpShape.height = desth
        
        If .tileset.selection(Button).selectionType = TS_Walltiles Then
        
            Call DrawImagePreview(picWalltiles, shpwallsel(Button), destpic, tmpShape, frmGeneral.TilesetBackgroundColor)
'                SetStretchBltMode pictilesetlarge.hDC, HALFTONE
'                StretchBlt pictilesetlarge.hDC, shppreviewsel(button).Left + 1, shppreviewsel(button).Top + 1, shppreviewsel(button).Width - 2, shppreviewsel(button).Height - 2, picWalltiles.hDC, (.tileset.selection(button).group Mod 4) * 64, (.tileset.selection(button).group \ 4) * 64, 4 * TILEW, 4 * TILEW, vbSrcCopy
'
'                'update the smalltile preview, used when the right panel is shrinked
'                BitBlt picsmalltilepreview.hDC, 1, 1 + (button - 1) * 18, TILEW, TILEW, pictileset.hDC, shptilesel(button).Left, shptilesel(button).Top, vbSrcCopy
'
        ElseIf .tileset.selection(Button).selectionType = TS_Tiles Then
            
            Call DrawImagePreview(pictileset, shptilesel(Button), destpic, tmpShape, frmGeneral.TilesetBackgroundColor)
        

'                SetStretchBltMode pictilesetlarge.hDC, COLORONCOLOR
'                StretchBlt pictilesetlarge.hDC, shppreviewsel(button).Left + 1, shppreviewsel(button).Top + 1, shppreviewsel(button).Width - 2, shppreviewsel(button).Height - 2, pictileset.hDC, shptilesel(button).Left, shptilesel(button).Top, TILEW, TILEW, vbSrcCopy
'
'                'update the smalltile preview, used when the right panel is shrinked
'                BitBlt picsmalltilepreview.hDC, 1, 1 + (button - 1) * 18, TILEW, TILEW, pictileset.hDC, shptilesel(button).Left, shptilesel(button).Top, vbSrcCopy
'
        ElseIf .tileset.selection(Button).selectionType = TS_LVZ Then
        
            Dim lvzidx As Integer
            Dim imgidx As Integer
            Dim srcDC As Long
            
            lvzidx = .tileset.selection(Button).group
            imgidx = .tileset.selection(Button).tilenr
            srcDC = .lvz.pichDClib(.lvz.getImageDefinition(lvzidx, imgidx).picboxIdx)
            
            
            Call DrawImagePreviewCoords(srcDC, 0, 0, .tileset.selection(Button).pixelSize.X, .tileset.selection(Button).pixelSize.Y, destpic.hDC, DestX, DestY, destw, desth, frmGeneral.TilesetBackgroundColor)
            
'                SetStretchBltMode pictilesetlarge.hDC, HALFTONE
'                StretchBlt pictilesetlarge.hDC, shppreviewsel(button).Left + 1, shppreviewsel(button).Top + 1, shppreviewsel(button).Width - 2, shppreviewsel(button).Height - 2, srcDC, 0, 0, .tileset.selection(button).pixelSize.X, .tileset.selection(button).pixelSize.Y, vbSrcCopy
'
'                StretchBlt picsmalltilepreview.hDC, 1, 1 + (button - 1) * 18, TILEW, TILEW, srcDC, 0, 0, .tileset.selection(button).pixelSize.X, .tileset.selection(button).pixelSize.Y, vbSrcCopy
'
'
        End If
        
        

    End With
    
'    pictilesetlarge.Refresh
'
''    'draw the 2 shapes that go around the selected tiles
''    Call DrawRectangle(picsmalltilepreview.hDC, 0, 0, 17, 17, vbRed)
''    Call DrawRectangle(picsmalltilepreview.hDC, 0, 18, 17, 36, vbYellow)
'    picsmalltilepreview.Refresh
End Sub
