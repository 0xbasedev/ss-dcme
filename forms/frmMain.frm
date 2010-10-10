VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "..."
   ClientHeight    =   8325
   ClientLeft      =   5460
   ClientTop       =   2115
   ClientWidth     =   11640
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   776
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox pictrans 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   6720
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Timer TimerAutosave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1560
      Top             =   7800
   End
   Begin VB.PictureBox picHighlightZoomTileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   11280
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox pichighlightTileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   10920
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox picseltemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15360
      Left            =   12960
      ScaleHeight     =   1024
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic1024selection 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15360
      Left            =   12000
      ScaleHeight     =   1024
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic1024 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15360
      Left            =   10560
      ScaleHeight     =   1024
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   10
      Top             =   6600
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox piczoomtileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   10560
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox pictileset 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   10560
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox picempty 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6240
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.VScrollBar vScr 
      Height          =   7575
      Left            =   10080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hScr 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7560
      Width           =   10095
   End
   Begin VB.PictureBox picpreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   513
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   680
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      Begin VB.Shape shpCorner 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000F&
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   10080
         Top             =   7560
         Width           =   255
      End
      Begin VB.Shape shptext 
         BorderColor     =   &H00FFFFFF&
         Height          =   135
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape shpcursor 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape shpdraw 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         Height          =   135
         Left            =   600
         Top             =   840
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum DrawLayers
    DL_Regions
    DL_LVZunder
    DL_Tiles
    DL_Selection
    DL_LVZover
    DL_Buffer
End Enum






Dim bmpFileheader As BITMAPFILEHEADER   'Holds the file header
Dim BMPInfoHeader As BITMAPINFOHEADER   'Holds the info header
Dim BMPRGBQuad() As RGBQUAD             'Holds the colour table
Dim bmpdata() As Byte                  'Holds the pixel data

'Color used as transparent key for the tiles layer so we
'can draw transparent fly over/under tiles
'By default this will be magenta
Dim unusedTilesetColor As Long


'pressing ctrl ?
Public usingctrl As Boolean

'pressing shift ?
Public usingshift As Boolean

'has map been changed ?
Public mapchanged As Boolean

'path to current map file, empty if unsaved
Public activeFile As String

'using the default tileset ? so no bitmap of tileset is present
Public usingDefaultTileset As Boolean

'holds the entire map
Dim tile(1023, 1023) As Integer

'different paste types
Public pastetype As enumPasteType

'grid on/off
Public usinggrid As Boolean
Public usinggridTest As Boolean 'grid for map testing

'regions shown on/off
Public ShowRegions As Boolean

'should regions be drawn (ShowRegions and regions.HaveVisibleRegions)
'updated at the start of UpdateLevel
Dim DrawRegions As Boolean

'lvz images shown on/off
Public ShowLVZ As Boolean

'should lvz's be drawn (ShowLVZ and lvz.HaveLVZ)
'updated at the start of UpdateLevel
Dim DrawLVZ As Boolean


'tilenr display on/off
Public usingtilenr As Boolean


'is the tileset already loaded?
Public tilesetloaded As Boolean

'store tileset's path
Public tilesetpath As String

'used for correct updating of the scrollbars (so they dont update both)
Dim dontUpdateOnValueChange As Boolean

'holds the current id of the map from the map
Public id As Integer

'holds the top and left of the radar
Public radar_left As Integer
Public radar_top As Integer

'is used to determine if we are using the dropper tool by holding ctrl
Public tempdropping As Boolean

'holds the tool previously used before the dropper tool was selected by ctrl
Dim toolbeforetempdropping As toolenum

'the from position where the selection began
Dim fromtilex As Integer
Dim fromtiley As Integer

Dim origX As Single
Dim origY As Single

'used to store the results of CheckCtrlShift
Dim tempx As Single
Dim tempy As Single

'used to remember last poly-line's size, for coordinates labels
Dim lastsplinesizex As Integer
Dim lastsplinesizey As Integer

Dim bookmark(0 To 9) As Long


'current zoom value (currentzoom)
Public currentzoom As Single
Public currenttilew As Integer 'precalculated tile width for current zoom

Dim regnopacity1 As Byte
Dim regnopacity2 As Byte

'contains information about currently selected tiles
Public WithEvents tileset As TilesetSelections
Attribute tileset.VB_VarHelpID = -1


'these are the tools
Public magnifier As magnifier
Public sel As selection
Public magicwand As magicwand

Public pencil As pencil
Public tline As Line
Public Dropper As Dropper
Public Bucket As Bucket
Public hand As hand
Public undoredo As UndoRedoStack
Public SPline As SPline
Public airbr As AirBrush
Public repbr As ReplaceBrush
Public TileText As TileText

Public TestMap As TestMap
Public walltiles As walltiles
Public eLVL As eLVLdata
Public Regions As Regions
Public RegionTool As RegionTool

Public Freehand As FreehandSelection
Public CFG As CfgSettings

Public lvz As LVZData

Dim MapLayers(DL_Regions To DL_Buffer) As clsDisplayLayer

Public TileRender As clsTileRender

'NOTE: once used all references to the hdc of these pictureboxes
'      need to go trough clsPic in order to function correctly!
'      That's why bitblt/stretchblt is also defined in clsPic !
'clsPic handles for fast pixel operations
Public cpic1024 As clsPic


'Grid Settings
Dim gridcolor(0 To 7) As Long

Dim GridOffsetX As Integer
Dim gridblocksX As Integer
Dim gridsectionsX As Integer

Dim GridOffsetY As Integer
Dim gridblocksY As Integer
Dim gridsectionsY As Integer

'Cursor settings
Private m_lcursorcolor As Long
Dim showtilepreview As Boolean
Dim showtilecoords As Boolean

'countdown for autosave
Dim MinutesCounted As Integer

Dim formReady As Boolean

Dim formActivated As Boolean






Sub Form_Activate()
'draw the tileset of the current map on the general form
    On Error GoTo Form_Activate_Error
    
'    If Not doneLoading Then Exit Sub
    picPreview.Enabled = True
    formActivated = True
    
'    Dim i As Integer
'    For i = DL_Regions To DL_Buffer
'        If MapLayers(i) Is Nothing Then
'            Set MapLayers(i) = New clsDisplayLayer
'        End If
'    Next
    
    If Not formReady Then Exit Sub
    
    
    
    Call Form_Resize
    
'    If PRINTDEBUG Then
'        AddDebug "At frmMain(" & id & ").Form_Activate"
'    End If
    
    Call frmGeneral.cTileset.CopyTilesetFromDC(pictileset.hDC, 0, 0)

    'make the activemap = this
    frmGeneral.activemap = id

    'now the activemap is set,
    'update the layout to match the attributes of the map
    frmGeneral.UpdateToolBarButtons

    'TOREMOVE --- REMOVED
    'set the tileset left and right
'    Call tileset.RefreshSelections
    Call frmGeneral.cTileset.setParent(Me)

    'select the this map under the menu items
    frmGeneral.UpdateMenuMaps

    'store settings
    Call StoreSettings

    If Not TestMap.isRunning Then
        'check if any other map is running testmap
        If frmGeneral.isTestmapActive(Me) Then
            'there is another one running
            'stop it
            Call frmGeneral.stopAllTestMap(Me)
        End If
    End If
        
    
    If Not frmGeneral.dontUpdatePreview Then
        'update the level map
        UpdateLevel
    End If
    frmGeneral.dontUpdatePreview = False

    Call DrawWallTilesPreview
'    Call frmGeneral.cTileset.Redraw
    
    frmGeneral.cTileset.DefaultLvzLayer = lvz.MapObjectDefaultLayer
    frmGeneral.cTileset.DefaultLvzMode = lvz.MapObjectDefaultMode
    
    
'    frmGeneral.tbTilesetView.Tabs(tileset.CurrentTab).selected = True
'
'    Call frmGeneral.tbTilesetView_Click
'
'
'    frmGeneral.cmbLVZTilesetDisplayType.ListIndex = lvz.MapObjectDefaultMode
'    frmGeneral.cmbLVZTilesetLayerType.ListIndex = lvz.MapObjectDefaultLayer
    
    
    'Call lvz.DrawLVZImageInterface(Rnd() * 1000)
    
    Call frmGeneral.UpdateToolBarButtons
    Call frmGeneral.UpdateToolToolbar

    
    On Error GoTo 0
    Exit Sub

Form_Activate_Error:
    HandleError Err, "frmMain(" & id & ").Form_Activate"
    
End Sub

Sub Form_Deactivate()
    shpcursor.visible = False
'    Call UpdatePreview(True, False)
    
'    Dim i As Integer
'    For i = DL_Regions To DL_Buffer
'        Set MapLayers(i) = Nothing
'    Next
    picPreview.Enabled = False
    
    formActivated = False
End Sub



Private Sub Form_Initialize()

    Dim i As Integer
    For i = DL_Regions To DL_Buffer
        Set MapLayers(i) = frmGeneral.GetLayer(i)
'        Call MapLayers(i).setParent(Me)
    Next
    
    
    
    Set tileset = New TilesetSelections
    
    
    'these are the tools
    Set magnifier = New magnifier
    Set sel = New selection
    Set magicwand = New magicwand

    Set pencil = New pencil
    Set tline = New Line
    Set Dropper = New Dropper
    Set Bucket = New Bucket
    Set hand = New hand
    Set undoredo = New UndoRedoStack
    Set SPline = New SPline
    Set airbr = New AirBrush
    Set repbr = New ReplaceBrush
    Set TileText = New TileText
    
    Set TestMap = New TestMap
    Set walltiles = New walltiles
    Set eLVL = New eLVLdata
    Set Regions = New Regions
    Set RegionTool = New RegionTool
    
    Set Freehand = New FreehandSelection
    Set CFG = New CfgSettings
    
    Set lvz = New LVZData
    
    Set TileRender = New clsTileRender
    
    
    'NOTE: once used all references to the hdc of these pictureboxes
    '      need to go trough clsPic in order to function correctly!
    '      That's why bitblt/stretchblt is also defined in clsPic !
    
    'clsPic handles for fast pixel operations
    Set cpic1024 = New clsPic

    
    


    Call tileset.setParent(Me)
    
    tileset.CurrentTab = TB_Tiles
    
    Call TileRender.setParent(Me)
    
'Initializes every tool used by the map
    Call undoredo.setParent(Me)
    Call sel.setParent(Me)
    Call magicwand.setParent(Me)
    Call magnifier.setParent(Me)

    Call pencil.setParent(Me)
    Call tline.setParent(Me)
    Call Dropper.setParent(Me)
    Call Bucket.setParent(Me)
    Call hand.setParent(Me)
    Call SPline.setParent(Me)
    Call airbr.setParent(Me)
    Call repbr.setParent(Me)
    Call walltiles.setParent(Me)
    Call TileText.setParent(Me)
    Call eLVL.setParent(Me)
    
    Call lvz.setParent(Me)
    
    
    Call Regions.setParent(Me)
    Call RegionTool.setParent(Me)
    Call CFG.setParent(Me)
    Call TestMap.setParent(Me)
    Call Freehand.setParent(Me)
    
    Call cpic1024.Init(pic1024)
    Call cpic1024.Clear
    
    formReady = True
    
'    Dim layer
'    For Each layer In MapLayers
''        Set layer = New clsDisplayLayer
'        Call layer.setParent(Me)
'    Next
    
'    doneInit = True
'
'    If doneLoad Then
'        doneLoading = True
'        Call Form_Resize
'        Call Form_Activate
'    End If
    
End Sub

Private Sub Form_Load()
'    If PRINTDEBUG Then
'10        subMain.AddDebug Me.hWnd & " Form loading"
'    End If
    
'Loads the default settings for a map
    Me.visible = False

    'use last grid settings
    usinggrid = CBool(GetSetting("ShowGrid", "1"))
    usinggridTest = CBool(GetSetting("ShowGridTest", "0")) 'show grid during map test
    
    usingtilenr = False
    
'Not used yet
'50        usinghidegrid = True

    ShowRegions = True
    ShowLVZ = True
    
    'normal pasting
    pastetype = p_normal

    shpcursor.Left = -20
    shpcursor.Top = -20

    'set default bookmarks
    Dim i As Integer
    For i = 0 To 9
        Call SetBookMark(i, GetBookMarkDefaultX(i), GetBookMarkDefaultY(i))
    Next

    
    'store settings
    Call StoreSettings

'    doneLoad = True
'

    Call Form_Activate


End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call frmGeneral.LoadMapFromOLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Close the map

    
    On Error GoTo Form_QueryUnload_Error
    
    If UnloadMode = QueryUnloadConstants.vbFormMDIForm Then
        'unload called because mdi is unloading
        'ignore
        Exit Sub
    ElseIf UnloadMode = QueryUnloadConstants.vbFormCode Then
        'unload called by direct code call, unload
        Call frmGeneral.DestroyMapReference(id)
        Unload Me
    Else


        'AddDebug "Unloading frmMain(" & id & ") " & Cancel & " " & UnloadMode
    
        If TestMap.isRunning Then Call TestMap.StopRun
    
        If mapchanged Then
            'the map has been changed, ask if the user wants to save it
            Dim retVal As Integer
            retVal = MessageBox(Me.Caption & " - " & "Not all changes have been saved, do you want to save first?", vbQuestion + vbYesNoCancel)
            If retVal = vbYes Then
                'don't close
                Dim saveok As Boolean
                saveok = frmGeneral.SaveMap(False, SFdefault)
    
                Cancel = Not saveok
                If Not saveok Then
                    Exit Sub
                End If
            ElseIf retVal = vbNo Then
                'no
                Cancel = False
            ElseIf retVal = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
    
    
        'ask the general form to destroy the map
        Cancel = False
        
        
        Call frmGeneral.DestroyMap(id)
        Call frmGeneral.UpdateMenuMaps

    End If
    
    On Error GoTo 0
    Exit Sub

Form_QueryUnload_Error:
    HandleError Err, "frmMain(" & id & ").Form_QueryUnload"
End Sub

Private Sub Form_Resize()
'Updates the controls when resizing
    On Error GoTo Form_Resize_Error
'    If PRINTDEBUG Then
'        AddDebug "At frmMain(" & id & ").Form_Resize (windowstate=" & Me.windowstate & ")"
'    End If
'          subMain.AddDebug Me.hWnd & " - Resizing form"
'    If Not doneLoading Then Exit Sub
    If Not formReady Then Exit Sub
    
    
    If frmGeneral.windowstate <> vbMinimized And Me.windowstate <> vbMinimized And formActivated Then
        'when the form is resized, the scrollbars and preview need to be
        'resized too
        On Error Resume Next
        'update scroll barrs
        hScr.Top = Me.ScaleHeight - hScr.height
        hScr.width = Me.ScaleWidth - vScr.width
        vScr.Left = Me.ScaleWidth - vScr.width
        vScr.height = Me.ScaleHeight - hScr.height
        
        shpCorner.Left = vScr.Left
        shpCorner.Top = hScr.Top
        
        'update preview picture
        Dim newWidth As Integer, newHeight As Integer
        
        newWidth = (Me.ScaleWidth - vScr.width) '+ currenttilew
        newHeight = (Me.ScaleHeight - hScr.height) ' + currenttilew
        
        'Minimum height
        If newHeight < 64 Then Me.height = (64 + hScr.height) * Screen.TwipsPerPixelY
        
'        'add a full tile
'        If newwidth Mod currenttilew <> 0 Then
'            newwidth = (newwidth \ currenttilew) * currenttilew + currenttilew
'        End If
'
'        If newheight Mod currenttilew <> 0 Then
'            newheight = (newheight \ currenttilew) * currenttilew + currenttilew
'        End If
'
'
'
'
        If TestMap.isRunning Then
            newWidth = newWidth + currenttilew
            newHeight = newHeight + currenttilew
        End If
        
        picPreview.width = newWidth
        picPreview.height = newHeight
        
'        If formActivated Then
            Dim i As Integer
            For i = DL_Regions To DL_Buffer
    '            If Not MapLayers(i).is_cached Then
                    Call MapLayers(i).Resize(newWidth, newHeight, True)
    '            End If
            Next
'        End If
        
        'because the form is resized, the preview can contain more data
        'from the level that was previously not loaded, so load it now
        Call UpdateScrollbars(True)
        
        
        Call UpdateLevel
    Else
        Call UpdateScrollbars(True)
    End If
    
'    If PRINTDEBUG Then
'        AddDebug "Finished frmMain(" & id & ").Form_Resize"
'    End If
    
    On Error GoTo 0
    Exit Sub
Form_Resize_Error:
    HandleError Err, "frmMain(" & id & ").Form_Resize"
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call tileset.setParent(Nothing)
    
'Initializes every tool used by the map
    Call undoredo.setParent(Nothing)
    Call sel.setParent(Nothing)
    Call magicwand.setParent(Nothing)
    Call magnifier.setParent(Nothing)
'    Call clip.setParent(Nothing)
    Call pencil.setParent(Nothing)
    Call tline.setParent(Nothing)
    Call Dropper.setParent(Nothing)
    Call Bucket.setParent(Nothing)
    Call hand.setParent(Nothing)
    Call SPline.setParent(Nothing)
    Call airbr.setParent(Nothing)
    Call repbr.setParent(Nothing)
    Call walltiles.setParent(Nothing)
    Call TileText.setParent(Nothing)
    Call eLVL.setParent(Nothing)
    Call lvz.setParent(Nothing)
    Call Regions.setParent(Nothing)
    Call RegionTool.setParent(Nothing)
    Call CFG.setParent(Nothing)
    Call TestMap.setParent(Nothing)
    Call Freehand.setParent(Nothing)
    
    Dim i As Integer
    For i = DL_Regions To DL_Buffer
'        Call MapLayers(i).setParent(Nothing)
        Set MapLayers(i) = Nothing
    Next
    
    Set tileset = Nothing
    
    
    'these are the tools
    Set magnifier = Nothing
    Set sel = Nothing
    Set magicwand = Nothing
'    Set clip = Nothing
    Set pencil = Nothing
    Set tline = Nothing
    Set Dropper = Nothing
    Set Bucket = Nothing
    Set hand = Nothing
    Set undoredo = Nothing
    Set SPline = Nothing
    Set airbr = Nothing
    Set repbr = Nothing
    Set TileText = Nothing
    
    Set TestMap = Nothing
    Set walltiles = Nothing
    Set eLVL = Nothing
    Set Regions = Nothing
    Set RegionTool = Nothing
    
    Set Freehand = Nothing
    Set CFG = Nothing
    
    Set lvz = Nothing
    
    Set TileRender = Nothing
    
    
    'NOTE: once used all references to the hdc of these pictureboxes
    '      need to go trough clsPic in order to function correctly!
    '      That's why bitblt/stretchblt is also defined in clsPic !
    
    'clsPic handles for fast pixel operations
    Set cpic1024 = Nothing
    
'    subMain.AddDebug "frmMain(" & id & ") unloaded"
End Sub

Private Sub Hscr_Change()
'Updates the horizontal scrollbar
'    If Hscr.value Mod currenttilew <> 0 Then
'        Hscr.value = (Hscr.value \ currenttilew) * currenttilew
'        Exit Sub
'    End If
    
'    frmGeneral.Label6.Caption = Hscr.value Mod currenttilew
    If Not dontUpdateOnValueChange Then
        UpdateLevel (True)
    End If

End Sub





Private Sub hScr_Scroll()
    Call Hscr_Change
    
    'put focus back on picpreview
    On Error Resume Next
    If Not TestMap.isRunning Then picPreview.setfocus
End Sub



Private Sub picPreview_KeyPress(KeyAscii As Integer)
    If curtool = T_tiletext Then
        Call TileText.KeyPress(KeyAscii)
    End If
End Sub

Private Sub picpreview_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call frmGeneral.LoadMapFromOLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub




Function hBufferDC() As Long
    hBufferDC = MapLayers(DL_Buffer).hDC
End Function



Private Sub tileset_OnChange()
    'Refresh the tileset display
    
    'TODO
    
End Sub




Private Sub TimerAutosave_Timer()
    'Counts a tick every minute

    MinutesCounted = MinutesCounted + 1
    
    If MinutesCounted >= CLng(GetSetting("AutoSaveDelay", DEFAULT_AUTOSAVE_DELAY)) And activeFile <> "" Then
        'call autosave
        If mapchanged Then
            frmGeneral.Label6.Caption = "Autosaving... " & MinutesCounted & " minutes elapsed."
        
            Call DoAutoSave
        
            MinutesCounted = 0
        End If
    End If
End Sub



Private Sub Vscr_Change()
'Updates the vertical scrollbar
'    If Vscr.value Mod currenttilew <> 0 Then
'        Vscr.value = (Vscr.value \ currenttilew) * currenttilew
'        Exit Sub
'    End If
    
    If Not dontUpdateOnValueChange Then
        UpdateLevel (True)
    End If


End Sub

Private Sub picpreview_KeyDown(KeyCode As Integer, Shift As Integer)
'When a key is pressed, use the shortcuts
    If TestMap.isRunning Then
        Exit Sub
    End If
    
    If curtool = T_tiletext Then
        Call TileText.KeyDown(KeyCode, Shift)
    Else
        If curtool = T_selection Or curtool = T_freehandselection Then
            If (Shift = 1 Or Shift = 2) Then
                picPreview.MousePointer = 2
            End If
        ElseIf curtool = T_magicwand Then
            If (Shift = 1 Or Shift = 2) Then
                picPreview.MousePointer = 99
            End If
        End If

        If KeyCode = vbKeyControl Then
            usingctrl = True
            
        ElseIf KeyCode = vbKeyShift Then
            usingshift = True
            
        ElseIf KeyCode = vbKeyEscape Then
            Select Case curtool
            Case T_selection
                If sel.selstate = drawing Then Call sel.StopSelecting
            End Select
        End If

        Call UseShortcutTool(True, KeyCode, Shift)
    End If
End Sub

Private Sub picpreview_KeyUp(KeyCode As Integer, Shift As Integer)
'When a key is pressed, use the shortcuts
    If TestMap.isRunning Then
        Exit Sub
    End If

    If curtool = T_tiletext Then
        Call TileText.KeyUp(KeyCode, Shift)
    Else
        If curtool = T_selection Then
            If (Shift = 1 Or Shift = 2) Then
                picPreview.MousePointer = 2
            End If
        ElseIf curtool = T_magicwand Then
            If (Shift = 1 Or Shift = 2) Then
                picPreview.MousePointer = 99
            End If
        End If

        If KeyCode = vbKeyControl Then
            usingctrl = False
        ElseIf KeyCode = vbKeyShift Then
            usingshift = False
        End If

        Call UseShortcutTool(False, KeyCode, Shift)
    End If
End Sub

Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Apply the current tool the map
'we have mousedown
    
    If SharedVar.MouseDown <> 0 Then Exit Sub 'mouse is already down
    
    SharedVar.MouseDown = Button

    Dim curtilex As Integer, curtiley As Integer




    'switch to hand with middle button
    If Button = vbLeftButton Then
        tileset.lastButton = vbLeftButton
    ElseIf Button = vbRightButton Then
          If Not SharedVar.splineInProgress Then
              tileset.lastButton = vbRightButton
          End If
    Else
        If Not tempdropping Then
            toolbeforetempdropping = curtool
        End If
        Call frmGeneral.SetCurrentTool(T_hand, False)
    End If

    OffsetToolCoordinate curtool, X, Y
    
    PlaceCursor X, Y

    origX = X
    origY = Y

    'calculate the tiles from the clicked pixel
    curtilex = (hScr.value + X) \ currenttilew
    curtiley = (vScr.value + Y) \ currenttilew
    If curtilex > 1023 Then curtilex = 1023
    If curtilex < 0 Then curtilex = 0
    If curtiley > 1023 Then curtiley = 1023
    If curtiley < 0 Then curtiley = 0
    
    
    'updates the fromtilex and fromtiley
    If Not SharedVar.splineInProgress Then
        fromtilex = curtilex
        fromtiley = curtiley
    End If



    Select Case curtool
    Case T_selection
        Call sel.MouseDown(Button, Shift, X, Y)

    Case T_magicwand
        Call magicwand.MouseDown(Button, Shift, X, Y)

    Case T_magnifier
        Call magnifier.MouseDown(Button, X, Y)

    Case T_hand
        Call hand.MouseDown(X, Y)

    Case T_pencil
        Call pencil.MouseDown(Button, X, Y)
        
    Case T_Eraser
        Call pencil.MouseDown(Button, X, Y)

    Case t_spline
        'make adjustments to make it perfectly diagonal
        'the results are stored in tempx and tempy
        Call CheckShiftCtrl(Shift, curtilex, curtiley)
        
        Call initCoordLabels(Button)
        'now use the enhanced coordinates (calculate them back to pixels)
        'in the mousedown
        Call SPline.MouseDown(Button, _
                              (tempx * currenttilew) - hScr.value, _
                              (tempy * currenttilew) - vScr.value)

        'update the from coordinates
        If SharedVar.splineInProgress Then
            fromtilex = tempx
            fromtiley = tempy
        End If

    Case T_dropper
        Call Dropper.MouseDown(Button, X, Y)

    Case T_bucket
        Call Bucket.MouseDown(Button, X, Y)
        
    Case T_line, T_rectangle, T_ellipse, T_filledellipse, T_filledrectangle, T_customshape
        Call tline.MouseDown(Button, X, Y)
            
    Case T_airbrush
        Call airbr.MouseDown(Button, X, Y)
        
    Case T_replacebrush
        Call repbr.MouseDown(Button, X, Y)
        
    Case T_tiletext
        Call TileText.MouseDown(Button, X, Y, Shift)

    Case T_Region
        Call RegionTool.MouseDown(Button, Shift, X, Y, frmGeneral.optRegionSel(REGION_MAGICWAND).value)

    Case T_lvz
        Call lvz.MouseDown(Button, Shift, X, Y)
        
    Case T_freehandselection
        Call Freehand.MouseDown(Button, Shift, X, Y)
        
    End Select

    If curtool <> t_spline Then
        Call initCoordLabels(Button)
    End If

End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Apply the tools when hovering over the preview

    Dim curtilex As Integer, curtiley As Integer
  
  Button = SharedVar.MouseDown
'10        If Button = vbLeftButton Then
'20            tileset.lastButton = vbLeftButton
'30        ElseIf Button = vbRightButton Then
'40            tileset.lastButton = vbRightButton
'50        End If

    If frmGeneral.AutoHideTileset Then
        Call frmGeneral.HideFloatTileset
    End If
    If frmGeneral.AutoHideRadar And frmGeneral.FloatRadar.autoHide Then
        Call frmGeneral.HideFloatRadar
    End If

    'Some icons have offsets from the left top of the icon, so
    'adjust the x and y values to those offsets
    OffsetToolCoordinate curtool, X, Y

    curtilex = (hScr.value + X) \ currenttilew
    curtiley = (vScr.value + Y) \ currenttilew

    If curtilex > 1023 Then curtilex = 1023
    If curtilex < 0 Then curtilex = 0
    If curtiley > 1023 Then curtiley = 1023
    If curtiley < 0 Then curtiley = 0

    'always show region name in tooltiptext, if there is a region there, and the coordinate

    
    Dim ret As Integer
'380       ret = Regions.isRegionAt(curtilex, curtiley)
    
    

    If Button = 0 Then Call SetTooltipText(X, Y, curtilex, curtiley)

    
    'update the position currently hovering over in tile coordinates
    If frmGeneral.lblposition.Caption <> "X= " & curtilex & " - Y= " & curtiley & " (" & Chr(65 + Int(curtilex / Int(1024 / 20))) & 1 + Int(curtiley / Int(1024 / 20)) & ")" Then
        'only update label when it has actually changed (to prevent flickering)
        frmGeneral.lblposition.Caption = "X= " & curtilex & " - Y= " & curtiley & " (" & Chr(65 + Int(curtilex / Int(1024 / 20))) & 1 + Int(curtiley / Int(1024 / 20)) & ")"
    End If

    PlaceCursor X, Y

    Select Case curtool
    Case T_selection
        'selection is also a rectangle so when pressing shift
        'make the selection square
        If sel.hasAlreadySelectedParts Then
            tempx = curtilex
            tempy = curtiley
        Else
            Call CheckShiftCtrl(Shift, curtilex, curtiley)
        End If
        '''''''''

        'now call the mousemove with possible enhanced coordinates
        Call sel.MouseMove(Button, Shift, (tempx * currenttilew) - hScr.value, _
                           (tempy * currenttilew) - vScr.value)

    Case T_magicwand
        Call magicwand.MouseMove(Button, Shift, X, Y)

    Case T_magnifier
        Call magnifier.MouseMove

    Case T_hand
        Call hand.MouseMove(Button, Shift, X, Y)

    Case T_pencil
        Call pencil.MouseMove(Button, X, Y)
    Case T_Eraser
        Call pencil.MouseMove(Button, X, Y)

    Case t_spline

        'calculate enhanced coordinates if we press shift or ctrl
        Call CheckShiftCtrl(Shift, curtilex, curtiley)

        'use the enhanced coordinates
        Call SPline.MouseMove(Button, _
                              (tempx * currenttilew) - hScr.value, _
                              (tempy * currenttilew) - vScr.value)

    Case T_dropper
        Call Dropper.MouseMove(Button, X, Y)

    Case T_bucket
        Call Bucket.MouseMove

    Case T_line, T_rectangle, T_ellipse, T_filledellipse, T_filledrectangle, T_customshape
        'calculate enhanced coordinates
        Call CheckShiftCtrl(Shift, curtilex, curtiley)

        'use those coordinates
        Call tline.MouseMove(Button, _
                             (tempx * currenttilew) - hScr.value, _
                             (tempy * currenttilew) - vScr.value, usingctrl)
    Case T_airbrush
        Call airbr.MouseMove(Button, X, Y)
    Case T_replacebrush
        Call repbr.MouseMove(Button, X, Y)
    Case T_tiletext
        Call TileText.MouseMove(Button, X, Y, Shift)
    Case T_Region
        Call RegionTool.MouseMove(Button, Shift, X, Y)
    Case T_lvz
        Call lvz.MouseMove(Button, Shift, X, Y)
    
    Case T_freehandselection
        Call Freehand.MouseMove(Button, Shift, X, Y)
    End Select

    Call updateCoordLabels(Button, curtilex, curtiley)

    'Move map while dragging if needed
    If Button And Not (curtool = T_hand Or curtool = T_airbrush Or curtool = T_bucket Or curtool = T_magnifier) Then
        Call DragMap(Button, X, Y, Shift)
    End If

End Sub

Private Sub picPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'do tool events when mouseup

'no more mouse down
  If Button <> SharedVar.MouseDown Then Exit Sub
  
    SharedVar.MouseDown = 0

    OffsetToolCoordinate curtool, X, Y
    

    Dim curtilex As Integer, curtiley As Integer
    'calculate tile coordinates
    curtilex = (hScr.value + X) \ currenttilew
    curtiley = (vScr.value + Y) \ currenttilew

      'Reset cursor color
    If Not showtilepreview Then shpcursor.BorderColor = cursorcolor

    Select Case curtool
    Case T_selection
        Call sel.MouseUp(Button, Shift, X, Y)

    Case T_magicwand
        Call magicwand.MouseUp(Button, Shift, X, Y)

    Case T_magnifier
        Call magnifier.MouseUp

    Case T_hand
        Call hand.MouseUp

    Case T_pencil, T_Eraser
        Call pencil.MouseUp

    Case t_spline
        Call SPline.MouseUp(Button, X, Y)

    Case T_bucket
        Call Bucket.MouseUp

    Case T_line, T_rectangle, T_ellipse, T_filledrectangle, T_filledellipse, T_customshape
        'calculate enhanced coordinates
        Call CheckShiftCtrl(Shift, curtilex, curtiley)
        'use mouseup with those coordinates
        Call tline.MouseUp(Button, _
                           (tempx * currenttilew) - hScr.value, _
                           (tempy * currenttilew) - vScr.value, _
                           usingctrl)
    Case T_airbrush
        Call airbr.MouseUp
    Case T_replacebrush
        Call repbr.MouseUp
    Case T_tiletext
        Call TileText.MouseUp(Button, X, Y, Shift)
    Case T_Region
        Call RegionTool.MouseUp(Button, Shift, X, Y)
    
    Case T_lvz
        Call lvz.MouseUp(Button, Shift, X, Y)
        
    Case T_freehandselection
        Call Freehand.MouseUp(Button, Shift, X, Y)
        
    End Select

    'if middle button goes up, switch back to previous tool
    If Button = vbLeftButton Then
    ElseIf Button = vbRightButton Then
    Else
        If Not tempdropping Then
            Call frmGeneral.SetCurrentTool(toolbeforetempdropping, False)
        Else
            Call frmGeneral.SetCurrentTool(T_dropper)
        End If
    End If


End Sub


Sub OffsetToolCoordinate(tool As toolenum, ByRef X As Single, ByRef Y As Single)
    Select Case tool
    Case T_pencil, T_replacebrush
        X = X - 7
        Y = Y + 7
    Case T_Eraser
        X = X - 7
        Y = Y + 5
    Case T_bucket
        X = X + 6
        Y = Y + 6
    Case T_dropper
        X = X - 7
        Y = Y + 7
    Case T_magicwand
        X = X + 7
        Y = Y - 8
    Case T_airbrush
        X = X - 9
        Y = Y + 6
    Case T_Region
        'same offset as magicwand with region tool in magic-wand mode
        If frmGeneral.optRegionSel(0).value Then
            X = X + 7
            Y = Y - 8
        End If
    End Select
    

    
End Sub

Private Sub SetTooltipText(X As Single, Y As Single, curtilex As Integer, curtiley As Integer)
    Dim newtooltiptext As String
    
    
    'Update tooltip text with tile coordinates
    If showtilecoords Then
        newtooltiptext = "(" & curtilex & " , " & curtiley & ")"
        
        Dim regns As String
        regns = Regions.getRegionsNamesAt(curtilex, curtiley)
        
        If regns <> "" Then newtooltiptext = newtooltiptext & " - " & regns
'410           If ret <> -1 Then
'420               newtooltiptext = newtooltiptext & " - " & Regions.getRegionName(ret)
'430           End If
    End If
    
    'Show LVZ information
    If ShowLVZ Then

        Dim lvzidx As Integer
        Dim moIdx As Long
        
        If lvz.getMapObjectAtPos((X + hScr.value) / magnifier.zoom, (Y + vScr.value) / magnifier.zoom, lvzidx, moIdx) <> -1 Then
            
            Dim mapobj As LVZMapObject
            mapobj = lvz.getMapObject(lvzidx, moIdx)
            
            newtooltiptext = newtooltiptext & "   " & "Layer: " & LVZLayerName(mapobj.layer) & _
                                        "   " & "Display: " & LVZModeName(mapobj.mode) & _
                                        "   " & "Display Time: " & mapobj.displayTime
        End If
    End If
    
    If picPreview.tooltiptext <> newtooltiptext Then
        picPreview.tooltiptext = newtooltiptext
    End If
End Sub


Sub ImportTileset(filename As String, Optional searchforwalltiles As Boolean = False)
'Imports a tileset into the current map
    Dim f As Integer
    On Error GoTo ImportTileset_Error

    f = FreeFile

    Dim bm As Integer
    Dim size As Long

    Dim b() As Byte

    Dim tmpfileh As BITMAPFILEHEADER
    Dim tmpinfoh As BITMAPINFOHEADER

    Open filename For Binary As #f
    ReDim b(1)
    Get #f, , b

    AddDebug "ImportTileset, importing: " & filename

    If Chr(b(0)) & Chr(b(1)) <> "BM" Then
        'if there is no BM it means that we have either a
        'corrupt lvl or bmp file, or we have a lvl file
        'that uses the default tileset
        AddDebug " --- No bitmap data found in " & filename

    Else
        'rewind to beginning
        Seek #f, 1
        Get #f, , tmpfileh
        Get #f, , tmpinfoh
        'check if its lower than 8 bit
        'and check for width=304 and height=160
        'if it's not, we have an invalid bmp or lvl
        'to import

        AddDebug BitmapHeaderInfoString(tmpinfoh)

        If tmpinfoh.biBitCount < 8 Or tmpinfoh.biHeight <> 160 Or tmpinfoh.biWidth <> 304 Then

            AddDebug "ImportTileset, invalid tileset, importing cancelled"

            MessageBox "Tilesets are required to be 304x160 pixels, with 8-bit or 24-bit color depth!", vbExclamation
            Close #f
            Exit Sub
        End If


        'ok valid tileset, importing
        bmpFileheader = tmpfileh
        BMPInfoHeader = tmpinfoh

        If tmpinfoh.biBitCount = 8 Then
            ReDim BMPRGBQuad(255)
            Get #f, , BMPRGBQuad
        End If

        'resize the data to fit the entire bmpdata
        'this data will only be used to save the map again
        If BMPInfoHeader.biCompression = 0 And BMPInfoHeader.biSizeImage = 0 Then
            'its possible the biSizeImage is 0
            If tmpinfoh.biBitCount = 8 Then
                ReDim bmpdata(BMPInfoHeader.biHeight * BMPInfoHeader.biWidth - 1)
            Else
                ReDim bmpdata(BMPInfoHeader.biHeight * BMPInfoHeader.biWidth * 3 - 1)
            End If
        Else
            'compression is used, read number of bytes from biSizeImage
            ReDim bmpdata(BMPInfoHeader.biSizeImage - 1)
        End If


        Get #f, , bmpdata

        AddDebug "ImportTileset, bmpdata is imported from: " & filename
        'this is the end of the tileset.
        'ok, we have the bmp data
    End If
    Close #f

    If searchforwalltiles Then
        Call SearchWallTiles(filename)
    End If

    'load the tileset from the lvl into the tileset and zoomtileset
    Call InitTileset(filename, True)

    'the map has changed
    mapchanged = True



    'update the menu's
    frmGeneral.mnudiscardtileset.Enabled = True
    frmGeneral.mnuExportTileset.Enabled = True

    'update the level to sync with the new tileset
    UpdateLevel

    On Error GoTo 0
    Exit Sub

ImportTileset_Error:
    HandleError Err, "frmMain" & id & ".ImportTileset " & filename
End Sub

Sub InitTileset(Optional path As String, Optional Refresh As Boolean = True)
'Load the tileset from the lvl into the tileset
'    On Error GoTo InitTileset_Error
    frmGeneral.IsBusy("frmMain" & id & ".InitTileset") = True
    
    Dim tick As Long
    
    tick = GetTickCount
    
    If path = "" Or path = "Default" Then
        'load it
        pictileset.Picture = frmGeneral.picdefaulttileset.Picture

        tilesetloaded = True
        tilesetpath = "Default"
        usingDefaultTileset = True
        unusedTilesetColor = RGB(255, 0, 255)
    Else
        'load the picture from the path
        pictileset.Picture = LoadPicture(path)
        tilesetloaded = True
        tilesetpath = path
        'not using default tileset anymore
        usingDefaultTileset = False
    End If

    AddDebug "InitTileset, tilesetpath= " & tilesetpath & " usingDefaultTileset " & usingDefaultTileset

    'create the correct size for the tileset
    pictileset.AutoSize = True
    pictileset.AutoSize = False
    pictileset.height = pictileset.height + frmGeneral.cTileset.Pic_ExtraTiles.height

    pictileset.Refresh

    'create the correct size for the zoomed tileset
    piczoomtileset.width = pictileset.width * currentzoom
    piczoomtileset.height = pictileset.height * currentzoom
    picHighlightZoomTileset.width = pictileset.width * currentzoom
    picHighlightZoomTileset.height = pictileset.height * currentzoom

    'bitblt the extratiles on the tileset, so they will be zoomed along with the others
    BitBlt pictileset.hDC, 0, pictileset.height - frmGeneral.cTileset.Pic_ExtraTiles.height, pictileset.width, frmGeneral.cTileset.Pic_ExtraTiles.height, frmGeneral.cTileset.Pic_ExtraTiles.hDC, 0, 0, vbSrcCopy
    Call DrawHighlightTileset(True, Not usingDefaultTileset)
    
    pictrans.BackColor = unusedTilesetColor
    pictrans.Cls
    
    Call RedrawGrid(False)
    Call RedrawTileLayer(True)
    
    
    
    pictileset.Refresh

    'draw the zoomed tiles onto the zoomtileset
    
    SetStretchBltMode piczoomtileset.hDC, COLORONCOLOR
    StretchBlt piczoomtileset.hDC, 0, 0, piczoomtileset.width, piczoomtileset.height, pictileset.hDC, 0, 0, pictileset.width, pictileset.height, vbSrcCopy
    piczoomtileset.Refresh

    SetStretchBltMode picHighlightZoomTileset.hDC, COLORONCOLOR
    StretchBlt picHighlightZoomTileset.hDC, 0, 0, piczoomtileset.width, piczoomtileset.height, pichighlightTileset.hDC, 0, 0, pichighlightTileset.width, pichighlightTileset.height, vbSrcCopy
    picHighlightZoomTileset.Refresh
    

    'copy the tileset data on the visible tileset
    Call frmGeneral.cTileset.CopyTilesetFromDC(pictileset.hDC, 0, 0)
    
    If Refresh Then frmGeneral.UpdatePreview

    Call frmGeneral.UpdateToolBarButtons

    Call DrawWallTilesPreview

    frmGeneral.IsBusy("frmMain" & id & ".InitTileset") = False
    
    
    frmGeneral.Label6.Caption = "Init: " & GetTickCount - tick
    On Error GoTo 0
    Exit Sub

InitTileset_Error:
    frmGeneral.IsBusy("frmMain" & id & ".InitTileset") = False
    HandleError Err, "frmMain" & id & ".InitTileset"
End Sub

Private Sub DrawHighlightTileset(highlightBlack As Boolean, findUnusedColor As Boolean)
    pichighlightTileset.width = pictileset.width
    pichighlightTileset.height = pictileset.height

    BitBlt pichighlightTileset.hDC, 0, 0, pictileset.width, pictileset.height, pictileset.hDC, 0, 0, vbSrcCopy

    Dim i As Integer
    Dim j As Integer
    
    Dim Pic As clsPic
    
    Set Pic = New clsPic
    
    Dim orgC As Long
    Dim newR As Byte, newG As Byte, newB As Byte
    
    Call Pic.Init(pichighlightTileset, True)
    
    Dim tick As Long
    tick = GetTickCount
    
    
    If findUnusedColor Then
        Dim c As Long
        c = 1
        Randomize
findAnotherColor:
        'Make sure it's not used in the grid
        For i = 0 To 7
            If c = gridcolor(i) Then
                c = Rnd() * RGB(255, 255, 255)
                GoTo findAnotherColor
            End If
        Next
        'Make sure it's not used in the tileset
        For j = 0 To pichighlightTileset.height - 1
            For i = 0 To pichighlightTileset.width - 1
                If Pic.GetPixel(i, j) = c Then
                    c = Rnd() * RGB(255, 255, 255)
                    GoTo findAnotherColor
                End If
            Next
        Next
        
        unusedTilesetColor = c
                
'        MsgBox GetTickCount - tick & " ms. Unused: " & unusedTilesetColor
    
    End If
    
    
    For j = 0 To pichighlightTileset.height - 1
        For i = 0 To pichighlightTileset.width - 1
            orgC = Pic.GetPixel(i, j)
            
     
            
            If orgC <> 0 Then
                newR = (GetRED(orgC) + 60) And &HFF
                newG = (GetGREEN(orgC) + 60) And &HFF
                newB = (GetBLUE(orgC) + 60) And &HFF
                
'                newR = IIf(GetRED(orgC) + 60 > 255, 255, GetRED(orgC) + 60)
'                newG = IIf(GetGREEN(orgC) + 60 > 255, 255, GetGREEN(orgC) + 60)
'                newB = IIf(GetBLUE(orgC) + 60 > 255, 255, GetBLUE(orgC) + 60)
                  Call Pic.SetPixel(i, j, newR, newG, newB)
                  
'110                   Call SetPixel(pichighlightTileset.hDC, i, j, RGB(newR, newG, newB))
              ElseIf highlightBlack Then
                  Call Pic.SetPixel(i, j, 60, 60, 60)
              End If
        Next
    Next

    Set Pic = Nothing
    
End Sub

Sub ExportTileset(path As String)
'Exports the current tileset
    On Error GoTo ExportTileset_Error

    If usingDefaultTileset Then
        'there is no tileset
        Exit Sub
    End If

    'if the user selected a file that already exists, delete the file
    DeleteFile path
'50        If FileExists(path) Then
'60            Kill path
'70        End If

    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    'open the file and write all bmp data stored
    Put #f, , bmpFileheader
    Put #f, , BMPInfoHeader

    If BMPInfoHeader.biBitCount = 8 Then
        Put #f, , BMPRGBQuad
    End If

    Put #f, , bmpdata

    AddDebug "ExportTileset, bmpdata is exported to: " & path
    AddDebug BitmapHeaderInfoString(BMPInfoHeader)

    Close #f

    On Error GoTo 0
    Exit Sub

ExportTileset_Error:
    HandleError Err, "frmMain" & id & ".ExportTileset " & path
End Sub

Sub DiscardTileset()
'Discard the current tileset
'ask for confirmation

    If MessageBox("Tileset will be lost, continue ?", vbYesNo + vbQuestion, "Discard tileset ?") = vbYes Then


        usingDefaultTileset = True
        AddDebug "DiscardTileset, usingDefaultTileset " & usingDefaultTileset
    
        'load the default tileset
        Call InitTileset("")
    
        'map has changed
        mapchanged = True
    
        'update level to sync with default tileset
        UpdateLevel
    End If

End Sub

Sub setTile(ByRef tileX As Integer, ByRef tileY As Integer, ByRef value As Integer, undoch As Changes, Optional appendundo = True)
    On Error GoTo setTile_Error

    'Sets the tile, and if required update the given undo changes stack
    If tile(tileX, tileY) <> value Then
        'only call undo when tile is changed
        Dim old As Integer
        old = tile(tileX, tileY)
        If appendundo Then
            Call undoch.AddTileChange(MapTileChange, tileX, tileY, old)
        End If
        tile(tileX, tileY) = value
    End If

    On Error GoTo 0
    Exit Sub
setTile_Error:
    'AddDebug "*** ERROR " & err.Number & ": " & err.Description & " *** @ setTile, map id " & id & " at (" & tileX & "," & tileY & ")"
    Resume Next
    '    messagebox "Error " & Err.Number & " (" & Err.Description & ") in procedure setTile of Form frmMain" & vbLf & "Your map has been saved to " & App.path & "\DCME_recovery.lvl.", vbCritical
    '    Call SaveMap(App.path & "\DCME_recovery.lvl")
    '    End
End Sub

Function getTile(ByRef tileX As Integer, ByRef tileY As Integer) As Integer
'Retrieves the value of a tile at the given coordinates
    getTile = tile(tileX, tileY)
End Function

Sub NewMap()
'Creates a new map with default tileset
    On Error GoTo NewMap_Error

    usingDefaultTileset = True
    AddDebug "NewMap, path = '' usingDefaultTileset " & usingDefaultTileset

    'use default walltiles
    Dim ret As String
    ret = GetSetting("DefaultWalltiles")

    AddDebug "NewMap, DefaultWalltiles = " & ret

    If FileExists(ret) Then
        AddDebug "NewMap, Loading walltiles... " & ret
        Call walltiles.LoadWallTiles(ret)
    End If

    Call walltiles.SetTileIsWalltile

    'use default tileset
    ret = GetSetting("DefaultTileset")

    AddDebug "NewMap, DefaultTileset = " & ret

    If FileExists(ret) Then
        Call ImportTileset(ret, False)
    Else
        Call InitTileset("", True)
    End If

    'Set the linewidth to 1


    dontUpdateOnValueChange = True
    Call UpdateScrollbars(False)
    'set the preview in the center of the map
    
    Call Form_Resize_Force
    
    
    Call SetFocusAt(512, 512, picPreview.width \ 2, picPreview.height \ 2, False)
    
    dontUpdateOnValueChange = False

    'the map hasn't been saved yet
    activeFile = ""

    Call frmGeneral.UpdateMenuMaps

    MinutesCounted = 0
    
    undoredo.ResetRedo
    
    
    Call UpdateLevel(False, True)
    
    'the map isn't changed as it has just been created
    mapchanged = False
    On Error GoTo 0
    Exit Sub

NewMap_Error:
    HandleError Err, "frmMain.NewMap"
End Sub

Sub OpenMap(filename As String)
'Opens a map
    On Error GoTo OpenMap_Error
    
    Dim undoch As Changes
    
    
'    Dim mapcorrupted As Boolean
'    mapcorrupted = False


  

    Dim b(3) As Byte
    
    Dim f As Integer
    f = FreeFile
    
    AddDebug "OpenMap, Opening Map... " & filename & " (" & f & ")"

    frmGeneral.IsBusy("frmMain" & id & ".OpenMap") = True
    
    'initialize autosave countdown
    MinutesCounted = 0
    
    Call frmGeneral.UpdateProgressLabel("Reading map header...")
    

    
    Open filename For Binary As #f
    Get #f, , b

    If Chr(b(0)) & Chr(b(1)) <> "BM" Then
        'there is no tileset attached to the map
        'so the default will be used
        'default on true
        usingDefaultTileset = True

        AddDebug "OpenMap, No tileset found ; usingDefaultTileset " & usingDefaultTileset

        'update menu
        frmGeneral.mnudiscardtileset.Enabled = False
        frmGeneral.mnuExportTileset.Enabled = False
        'resize magnifier.zoom tileset
        'reset position
        Seek #f, 1
        'Now we are at the correct position for the tiles
    Else
        'update the menu's
        frmGeneral.mnudiscardtileset.Enabled = True
        frmGeneral.mnuExportTileset.Enabled = True
        'rewind to beginning
        Seek #f, 1
        Get #f, , bmpFileheader
        Get #f, , BMPInfoHeader
        'check if its 8/24 bit, normally it would be else we
        'have a corrupt lvl file
        'but check anyway
        AddDebug "OpenMap, Tileset found"
        AddDebug "OpenMap, " & BitmapHeaderInfoString(BMPInfoHeader)
        AddDebug "OpenMap, " & BitmapFileInfoString(bmpFileheader)

        If BMPInfoHeader.biBitCount < 8 Then
            AddDebug "OpenMap, Tileset is invalid"
            MessageBox "Invalid tileset within lvl file!", vbExclamation
            Close #f
            frmGeneral.IsBusy("frmMain" & id & ".OpenMap") = False
            Exit Sub
        End If

        If BMPInfoHeader.biBitCount = 8 Then
            ReDim BMPRGBQuad(255)
            Get #f, , BMPRGBQuad
        Else
            'skip because everything bigger than 8-bit depth
            'wont be using a rgbquad anymore
        End If

        'resize bmpdata
        If BMPInfoHeader.biCompression = 0 And BMPInfoHeader.biSizeImage = 0 Then
            'its possible the biSizeImage is 0
            If BMPInfoHeader.biBitCount = 8 Then
                ReDim bmpdata(BMPInfoHeader.biHeight * BMPInfoHeader.biWidth - 1)
            Else
                ReDim bmpdata(BMPInfoHeader.biHeight * BMPInfoHeader.biWidth * 3 - 1)
            End If
        Else
            'compression is used, read number of bytes from biSizeImage
            ReDim bmpdata(BMPInfoHeader.biSizeImage - 1)
        End If

        Get #f, , bmpdata

        AddDebug "Openmap, BMPData is read from lvl file"


        'check if eLVL data is present
        
        If LongToUnsigned(bmpFileheader.bfReserved1) > 0 Then
            'we most likely have eLVL data
            
            AddDebug "OpenMap, trying to read eLVL data"

            eLVL.GetELVLData f, LongToUnsigned(bmpFileheader.bfReserved1), filename
            
            'now seek to the part where the map bytes are stored
            'Seek #f, (BMPFileHeader.bfSize + 1)
            
        Else
            'probably no eLVL data in that map
            AddDebug "OpenMap, no eLVL data found"
            
            'we should already be at the correct place in the file to load tiles
        End If


        'this is the end of the tileset, tiles should start here

        usingDefaultTileset = False
        AddDebug "Openmap, usingDefaultTileset " & usingDefaultTileset
    End If
      
'          On Local Error GoTo Corrupt_File
    
    'we should already be here
    If Seek(f) <> bmpFileheader.bfSize + 1 Then
        If bmpFileheader.bfSize = 49720 Then
            Seek #f, bmpFileheader.bfSize + 1
        End If
        
'        AddDebug "OpenMap, WARNING: Seek(f)=" & Seek(f) & " bfSize+1=" & BMPFileHeader.bfSize + 1
'
'        Dim startseek As Long
'        startseek = Seek(f)
'        'Stop
'        'something is wrong here, attempt to recovery file
'        Dim recoverresult As Long
'
'        recoverresult = TileDataRecovery(longMinimum(BMPFileHeader.bfSize + 1, startseek), f, 4)
''        recoverresult = TileDataRecovery(BMPFileHeader.bfSize + 1, f, 4)
'        If recoverresult < 0 Then
'            recoverresult = TileDataRecovery(startseek, f, 4)
'        End If
'        If recoverresult < 0 Then
'            recoverresult = TileDataRecovery(longMinimum(BMPFileHeader.bfSize + 1, startseek), f, longMaximum(Abs((BMPFileHeader.bfSize + 1) - startseek), 30))
'        End If
'
'        Seek #f, Abs(recoverresult)

    End If
    
    Call frmGeneral.UpdateProgress("Loading map", 100)
    
    'search for walltiles, if exists and loaded , refresh
    'SearchWallTiles (FileName)
    'NOT NEEDED ANYMORE: Walltiles are in eLVL

    Call walltiles.SetTileIsWalltile
    'End If

    'update the tileset
    If usingDefaultTileset Then
        Call InitTileset("")
    Else
        Call frmGeneral.UpdateProgressLabel("Initializing tileset...")
        
        Call InitTileset(filename)
    End If




    AddDebug "OpenMap, tile data starting at " & Seek(f)
    
    Call frmGeneral.UpdateProgress("Loading map", 200)
    
    Call frmGeneral.UpdateProgressLabel("Loading tiles...")
    
    Dim nrtiles As Long
    nrtiles = 0
    
    
    Dim loadcorrupttiles As Boolean, askedloadcorrupt As Boolean
    Dim corruptcount As Long
    
    
    'keep em coming till we have reached the end of the file
    Do Until EOF(f)
        Dim X As Integer, Y As Integer

        'retrieve 4 bytes
        Get #f, , b

        'extract the data
        X = (b(0) + 256 * (b(1) Mod 16)) 'Mod 1024
        Y = (b(1) \ 16 + 16 * b(2)) 'Mod 1024
        
        
        If X < 0 Or X > 1023 Or Y < 0 Or Y > 1023 Then
            'We have a corrupt tile
            If Not askedloadcorrupt Then
                If MessageBox("Some tiles appear to be corrupted, do you want to load them anyway?", vbYesNo + vbQuestion) = vbYes Then
                    loadcorrupttiles = True
                End If
                askedloadcorrupt = True
            End If
            
            If loadcorrupttiles Then
                X = X Mod 1024
                Y = Y Mod 1024
            End If
            
            corruptcount = corruptcount + 1
        End If
        
        If X >= 0 And X <= 1023 And Y >= 0 And Y <= 1023 Then
        
            If tile(X, Y) > 0 Then
                'tile has already come, skip it, prolly last tile
            ElseIf b(3) <> 0 Then
                'last entry might be 0, just don't count it
                
                tile(X, Y) = b(3)
                nrtiles = nrtiles + 1
                
                Call CompleteObject(Me, X, Y, undoch, False)
                
            End If
        End If
    Loop
    
    Call frmGeneral.UpdateProgress("Loading map...", 600)
    
    AddDebug "OpenMap, " & nrtiles & " tiles loaded. Now at " & Seek(f)
    
    
    Close #f

    If corruptcount > 0 Then
        AddDebug corruptcount & " corrupted tiles " & IIf(loadcorrupttiles, "loaded", "ignored")
        
        MessageBox "A total of " & corruptcount & " corrupted tiles were " & IIf(loadcorrupttiles, "loaded.", "ignored."), vbInformation + vbOKOnly
    End If
    
    'we load so we have a path
    activeFile = filename
    
    AddDebug "OpenMap, File closed. Searching for matching cfg file"
    Call CFG.SearchCfg
    
    'add this to recent
    AddDebug "OpenMap, Adding to recent list"
    Call frmGeneral.AddRecent(filename)

    'just opened so the map hasn't changed yet
    mapchanged = False


    dontUpdateOnValueChange = True
    Call UpdateScrollbars(False)


    dontUpdateOnValueChange = False

    AddDebug "OpenMap, Drawing tiles"
    
    Call frmGeneral.UpdateProgressLabel("Drawing tiles...")
    
    'draw the full map used for radar or usingpixels magnifier.zoom level
    Dim i As Integer
    Dim j As Integer
    For j = 0 To 1023
        For i = 0 To 1023
            If tile(i, j) <> 0 Then
                'when an object is special, flag the other
                'tiles to -1 so that we know we are in an object
                

                'draw the complete pixel map
                'Call setPixel(pic1024.hdc, i, j, TilePixelColor(tile(i, j)))
                Call cpic1024.setPixelLong(i, j, TilePixelColor(tile(i, j)))
            End If
        Next
        
        If j Mod 4 = 0 Then Call frmGeneral.UpdateProgress("Loading map", 601 + j)
    Next
    
'1140      AddDebug "OpenMap, Saving revert copy"
'1150      Call SaveRevert
    
    
    
    Call Form_Resize

    
    Call Regions.RedrawAllRegions
    
    'same as with new map, go to center of the map
    Call SetFocusAt(512, 512, picPreview.width \ 2, picPreview.height \ 2, False)
    
    undoredo.ResetRedo
    Call frmGeneral.UpdateMenuMaps

    mdlDebug.AddDebug "Map loaded successfully"
    
    Call UpdateLevel(False, True)
    
    frmGeneral.IsBusy("frmMain" & id & ".OpenMap") = False
    
    
    On Error GoTo 0
    Exit Sub


'Corrupt_File:
'    If mapcorrupted Then
'        'Error appeared more than once, ignore
'        Resume Next
'    Else
'        AddDebug "OpenMap, Map corrupted! Seek position: " & Seek(f) & " last (X,Y): (" & X & "," & Y & ") b(3)= [" & b(0) & "," & b(1) & "," & b(2) & "," & b(3) & "]"
'        mapcorrupted = True
'        If MessageBox(filename & " may be corrupted. Do you wish to attempt recovering data?", vbYesNo + vbCritical) = vbYes Then
'            Dim recovery As Long
'            recovery = TileDataRecovery(BMPFileHeader.bfSize + 1, f)
'            If recovery < 0 Then
'                If MessageBox(filename & " could not be recovered completly, do you still want to load the file?", vbQuestion + vbYesNo) = vbYes Then
'                    Seek #f, Abs(recovery)
'                Else
'                    GoTo AbortRecovery
'                End If
'            Else
'                Seek #f, Abs(recovery)
'            End If
'
'            Resume Next
'        Else
'            GoTo AbortRecovery
'        End If
'    End If
'
'    frmGeneral.IsBusy("frmMain" & id & ".OpenMap") = False
'    Exit Sub
'
'AbortRecovery:
'    Close #f
'    frmGeneral.IsBusy("frmMain" & id & ".OpenMap") = False
'    Exit Sub
    
OpenMap_Error:
    frmGeneral.IsBusy("frmMain" & id & ".OpenMap") = False
    HandleError Err, "frmMain.OpenMap " & filename
End Sub

Sub SaveMap(path As String, Optional flags As saveFlags = SFdefault)
'Saves the map
'back the old one up
'if there is already a backup, kill it
'then rename the new old one to the .bak one
    On Error GoTo SaveMap_Error

    AddDebug "SaveMap, saving map... " & path
    AddDebug "         Flags = " & flags
    
    If FlagIs(flags, SFsilent) Then
        frmGeneral.SetMousePointer vbArrowHourglass
    Else
        
        frmGeneral.IsBusy("frmMain" & id & ".SaveMap") = True
        
        If FileExists(path) Then
            AddDebug "SaveMap, " & path & " exists, checking for .bak file"
            If DeleteFile(path & ".bak") Then
                  RenameFile path, path & ".bak"
            End If
'100               If FileExists(path & ".bak") Then
'110                   AddDebug "SaveMap, " & path & ".bak exists, killing it"
'120                   Kill path & ".bak"
'130               End If
'140               Name path As path & ".bak"
          
            'Make sure the file is deleted
            DeleteFile path
        End If
    
    End If
    


    Dim f As Integer
    f = FreeFile
    
    AddDebug "SaveMap, opening " & path & " for binary as #" & f
    
    
    
    Open path For Binary As #f
    
    On Error GoTo SaveMap_Error_Opened
    
    '''''TILESET IS NOW ALWAYS SAVED IF WE HAVE ELVL DATA


        
          
    If (usingDefaultTileset And Not (eLVL.HasData Or lvz.getLVZCount > 0)) Or Not FlagIs(flags, SFsaveTileset) Then
        AddDebug "SaveMap, usingDefaultTileset ; no tileset saved"
        'no need to write the bmp
        'just begin writing the tile data
    Else
        If usingDefaultTileset Then
            'We'll need to import the default tileset as a new tileset
            frmGeneral.picdefaulttileset.Picture = frmGeneral.picdefaulttileset.Image
            
            Dim temppath As String
            temppath = App.path & "\tempTileset" & GetTickCount & ".bmp"
            Call SavePicture(frmGeneral.picdefaulttileset.Picture, temppath)
            
            Call ImportTileset(temppath)
            
            DeleteFile temppath
        End If
        
        Call frmGeneral.UpdateProgressLabel("Saving tileset...")
        
        'TODO: IF SAVEASSSMECOMPATIBLE, CONVERT TO 8 BIT
        '##################
        '##################
        '##################
        
        'put the bitmap data first

        'bitmap header position is needed to replace it later when we know the size of eLVL data
        Dim headerPos As Long
        headerPos = Seek(f)
        AddDebug "SaveMap, bitmap header position: " & headerPos

        Put #f, , bmpFileheader
        Put #f, , BMPInfoHeader

        If BMPInfoHeader.biBitCount = 8 Then
            Put #f, , BMPRGBQuad
        End If

        Put #f, , bmpdata
    End If

    Call frmGeneral.UpdateProgress("Saving", 50)
    

    
    If eLVL.HasData And FlagIsPartial(flags, SFsaveELVL) Then

        'NOT TRUE ANYMORE >>>>>>>>>>>>>
        'Can't save eLVL with 24bit tileset
        'If BMPInfoHeader.biBitCount = 24 Then
        '    messagebox "eLVL data could not be saved because the tileset has a color depth of 24 bits. If you wish to save regions and attributes, you will need to convert your tileset to 8 bits using any drawing program, re-import it, and save the map again. You map will be saved without eLVL data for now.", vbOKOnly + vbExclamation, "eLVL data not saved"
        '<<<<<<<<<<<<<<<<<<<<<<

        'Can't save eLVL without tileset

        If usingDefaultTileset Or Not FlagIs(flags, SFsaveTileset) Then
            AddDebug "SaveMap, No tileset found in map, cannot save eLVL data"
            If Not FlagIs(flags, SFsilent) Then
                MessageBox "eLVL data could not be saved because this map does not have a tileset. If you wish to save regions and attributes, you will need to to import a tileset, and save the map again. You map will be saved without eLVL data for now.", vbOKOnly + vbExclamation, "eLVL data not saved"
            End If

        Else
            Call frmGeneral.UpdateProgressLabel("Saving eLVL data...")
        
            ''''put eLVL data
'            If True Then  'UnsignedToInteger(eLVL.Next4bytes(Seek(f)) - 1) <= 32767
            bmpFileheader.bfReserved1 = UnsignedToLong(Next4bytes(Seek(f)) - 1)    'should be 49720 with 8bit tilesets, unless compression
'            Else
'                AddDebug "SaveMap, bfReserved1 could not be set to " & (eLVL.Next4bytes(Seek(f)) - 1)
'                messagebox "map not saved, view debug for more info"    'TEMPORARY ERROR MESSAGE
'                Close #f
'                Exit Sub
'            End If

            Seek #f, Next4bytes(Seek(f))

            AddDebug "SaveMap, bfReserved1 set to: " & bmpFileheader.bfReserved1 & " (" & LongToUnsigned(bmpFileheader.bfReserved1) & ")"

            'Seek #f, IntegerToUnsigned(BMPFileHeader.bfReserved1) + 1
            Dim elvlsize As Long, orgbfSize As Long
            
            orgbfSize = bmpFileheader.bfSize
            

            elvlsize = eLVL.PutELVLData(f, flags)

            AddDebug "SaveMap, total eLVL size returned: " & elvlsize

            'update bitmap filesize including elvl data
            bmpFileheader.bfSize = UnsignedToLong((LongToUnsigned(bmpFileheader.bfReserved1) + elvlsize))

            AddDebug "SaveMap, New bitmap bfSize: " & bmpFileheader.bfSize & " (" & LongToUnsigned(bmpFileheader.bfSize) & ")"

            ''''replace bitmap header with new values
            Seek #f, headerPos
            Put #f, , bmpFileheader

            'once all the bmp data and elvl is put, continue at the position for
            'the map data
            Seek #f, LongToUnsigned(bmpFileheader.bfSize) + 1
            
            'Reset header properties
            bmpFileheader.bfReserved1 = 0
            bmpFileheader.bfSize = orgbfSize
        End If

    End If


    
    Call frmGeneral.UpdateProgress("Saving", 100)
    Call frmGeneral.UpdateProgressLabel("Saving tiles...")
        
    AddDebug "SaveMap, BMPData is stored into lvl, starting tile data at " & Seek(f)


    Dim nrtiles As Long
    nrtiles = 0

    Dim X As Integer, Y As Integer, curtile As Integer
    Dim Data As Long, btile As Byte
    
    Dim dontSaveExtraTiles As Boolean
    
'    Dim tick As Long
'    tick = GetTickCount
    
    dontSaveExtraTiles = Not FlagIs(flags, SFsaveExtraTiles)
    
    For Y = 0 To 1023

        
        For X = 0 To 1023
        
            curtile = tile(X, Y)
            
            'watch out for special tile flagging !, but
            'those will be 0 anyway, so skip them too
            If curtile > 0 Then
                  
                
                'if ssme compatible is required, discard all tiles
                'above 190 except for the ones that ssme does recognize
                If dontSaveExtraTiles Then
                    If curtile > 190 And _
                        curtile <> 216 And curtile <> 217 And _
                        curtile <> 218 And curtile <> 219 And _
                        curtile <> 220 Then
                        
                        'do not save those not compatible with ssme
                        
                        
                    Else
                        'now put the bytes
                        
                        Data = (X Mod 256) + (Y \ 16) * &H10000 + (X \ 256) * &H100 + &H1000 * CLng(Y Mod 16)
                        
                        btile = curtile
                        
                        CopyMemory ByVal VarPtr(Data) + 3, ByVal VarPtr(btile), 1
                        
                        Put #f, , Data

'                        Put #f, , CByte(X Mod 256)
'                        Put #f, , CByte((X \ 256) + (Y Mod 16) * 16)
'                        Put #f, , CByte(Y \ 16)
'
'                        Put #f, , CByte(curtile)

    
                        nrtiles = nrtiles + 1
                    End If
                Else
                
'
                    
                    
                    'Doing it this way only uses 1 'Put' instruction instead of 4, which reduces
                    'by almost 4x the execution time
                    'We have to use a CopyMemory for the MSB of the Long because otherwise we'd get
                    'overflow errors
                    Data = (X Mod 256) + (Y \ 16) * &H10000 + (X \ 256) * &H100 + &H1000 * CLng(Y Mod 16)
                    
                    btile = curtile
                    
                    CopyMemory ByVal VarPtr(Data) + 3, ByVal VarPtr(btile), 1
                    
                    Put #f, , Data
                    
                    'now put the bytes
'                    Put #f, , CByte(X Mod 256)
'                    Put #f, , CByte((X \ 256) + (Y Mod 16) * 16)
'                    Put #f, , CByte(Y \ 16)
'                    Put #f, , CByte(curtile)

                    
                    nrtiles = nrtiles + 1
                End If
            'Else
                'dont save those empty tiles
            End If
        Next
        
'        If Y Mod 4 = 0 Then Call frmGeneral.UpdateProgress("Saving", 101 + Y)

    Next
    
    
    AddDebug "SaveMap, " & nrtiles & " tiles were saved into lvl. Total file size: " & Seek(f) & " bytes."
      
    AddDebug "SaveMap, Closing file #" & f
    Close #f

    
    'save lvzs
    If FlagIs(flags, SFsaveLVZ) And lvz.getLVZCount > 0 Then
        Call frmGeneral.UpdateProgressLabel("Saving LVZ files...")
        
        Dim i As Integer
        For i = 0 To lvz.getLVZCount - 1
            
            Call lvz.exportLVZ(GetPathTo(path) & lvz.getLVZname(i), i)
            
            Call frmGeneral.UpdateProgress("Saving", 1124 + (200 / lvz.getLVZCount) * (i + 1))
        Next
    End If
    
    Call frmGeneral.UpdateProgress("Saving", 1324)
    
    'reset autosave countdown
    MinutesCounted = 0
    

        
    If FlagIs(flags, SFsaveExtraTiles) And Not FlagIs(flags, SFsilent) Then
        AddDebug "mapchanged set to False"
        mapchanged = False
        'else, no mapchanged, as some data could be lost
    End If
        
    Call frmGeneral.UpdateMenuMaps
    
'930       If FlagIs(flags, SFsaveRevert) Then
''940           Call frmGeneral.UpdateProgressLabel("Creating backup for revert...")
''950           Call SaveRevert
'960       End If
    
    Call frmGeneral.UpdateProgress("Saving", 1524)
    
    frmGeneral.IsBusy("frmMain" & id & ".SaveMap") = False
    

    
    On Error GoTo 0
    Exit Sub

SaveMap_Error_Opened:
    Close #f
SaveMap_Error:

    frmGeneral.IsBusy("frmMain" & id & ".SaveMap") = False

    If Err.Number = 75 Then
        MessageBox "Path/File access error 75" & vbCrLf & "You do not have permissions to write to '" & path & "'. Make sure you have the permission to write to this folder.", vbExclamation + vbOKOnly
        Exit Sub
    Else
        HandleError Err, "frmMain" & id & ".SaveMap " & path & " " & flags & "(" & X & "," & Y & ")", Not FlagIs(flags, SFsilent)
    End If
End Sub

Sub SwitchOrReplace(src As Integer, dest As Integer, replace As Boolean, Redraw As Boolean)
'Switches or replaces tiles on the entire map

    Dim i As Integer
    Dim j As Integer

    '2 counters that hold how many tiles are replaced or switched
    Dim n As Long
    Dim m As Long
    Dim p As Long
    'prepare the undoch
    On Error GoTo SwitchOrReplace_Error

    undoredo.ResetRedo
    Dim undoch As Changes
    Set undoch = New Changes

    n = 0
    m = 0
    p = 0

    Dim issrcvalid(255) As Boolean
    Dim isdestvalid(255) As Boolean
    Dim srcwalltile As Boolean
    Dim destwalltile As Boolean
    Dim srcwallset As Integer
    Dim destwallset As Integer
    Dim tmp As Integer
    Dim curtile As Integer

    If src >= 259 Then
        srcwallset = src - 259
        srcwalltile = True
        For i = 0 To 15
            issrcvalid(walltiles.getWallTile(src - 259, i)) = True
        Next
        src = walltiles.getWallTile(srcwallset, 0)
    End If

    If dest >= 259 Then
        destwallset = dest - 259
        destwalltile = True
        For i = 0 To 15
            isdestvalid(walltiles.getWallTile(dest - 259, i)) = True
        Next
        dest = walltiles.getWallTile(destwallset, 0)
    End If

    For j = 0 To 1023
        For i = 0 To 1023
            'replace the src tile with the dest tile
            If tile(i, j) < 0 Then
                curtile = tile(i, j) \ -100
            Else
                curtile = tile(i, j)
            End If

            If curtile = src Or (srcwalltile And issrcvalid(curtile)) Then
                If destwalltile Then
                    tmp = walltiles.ReplaceWithWalltile(i, j, destwallset, src, srcwalltile, srcwallset, replace)
                Else
                    tmp = dest
                End If

                Call setTile(i, j, tmp, undoch)
                Call UpdateLevelTile(i, j, False)
                n = n + 1
                'if the tile is a dest tile and we switch
                'then we replace it with the src tile
            ElseIf Not replace And (curtile = dest Or (destwalltile And isdestvalid(curtile))) Then
                If srcwalltile Then
                    tmp = walltiles.ReplaceWithWalltile(i, j, srcwallset, dest, destwalltile, destwallset, replace)
                Else
                    tmp = src
                End If

                Call setTile(i, j, tmp, undoch)
                Call UpdateLevelTile(i, j, False)
                m = m + 1

            ElseIf replace And Redraw And destwalltile And isdestvalid(curtile) Then

                tmp = walltiles.ReplaceWithWalltile(i, j, destwallset, src, srcwalltile, srcwallset, True)

                Call setTile(i, j, tmp, undoch)
                Call UpdateLevelTile(i, j, False)
                p = p + 1

            End If
        Next
    Next

    'show the results of the operation
    If replace Then
        MessageBox n & " tiles replaced" & vbCrLf & p & " walltiles redrawn", vbOKOnly + vbInformation, "Tiles replaced"
    Else
        MessageBox n & " tiles A -> B" & vbCrLf & m & " tiles A <- B", vbOKOnly + vbInformation, "Tiles switched"
    End If

    'add the changes to the undo stack
    Call undoredo.AddToUndo(undoch, IIf(replace, UNDO_REPLACE, UNDO_SWITCH))

    'update the level
    Call UpdateLevel

    'map has changed
    'mapchanged = True

    On Error GoTo 0
    Exit Sub

SwitchOrReplace_Error:
    HandleError Err, "frmMain.SwitchOrReplace"
End Sub

Function CountTile(tilenr As Integer) As Long
'Count the number of flags in the map
    Dim count As Long
    Dim i As Integer
    Dim j As Integer

    count = 0
    For j = 0 To 1023
        For i = 0 To 1023
            'tile 170 = flag
            If tile(i, j) = tilenr Then
                count = count + 1
            End If
        Next
    Next

    CountTile = count
End Function


Function CountTiles(inselection As Boolean, countPtr As Long) As Long
'Counts the tiles on the map (or in selection)
'Returns the total number of tiles in the map

    On Error GoTo CountTiles_Error
    
    Dim i As Integer, j As Integer

    'Will hold each tile's count
    Dim tilescount(255) As Long
    
    Dim total As Long
    
      
    'holds the boundaries to scan in
    Dim lbx As Integer
    Dim ubx As Integer
    Dim lby As Integer
    Dim uby As Integer
    

    'put the correct boundaries
    If inselection Then
        lbx = sel.getBoundaries.Left
        ubx = sel.getBoundaries.Right
        lby = sel.getBoundaries.Top
        uby = sel.getBoundaries.Bottom
    Else
        lbx = 0
        ubx = 1023
        lby = 0
        uby = 1023
    End If

    Dim curtile As Integer
    
    'count the tiles in the area, when a special tile is selected
    'count it once with n or m and once with special
    For j = lby To uby
        For i = lbx To ubx
            If inselection Then
                If sel.getIsInSelection(i, j) Then
                      curtile = sel.getSelTile(i, j)
                Else
                      curtile = -1
                End If
            Else
              curtile = tile(i, j)
            End If
            
            If curtile > 0 Then
              tilescount(curtile) = tilescount(curtile) + 1
              total = total + 1
            End If
        Next
    Next

    'copy the array
    
    CopyMemory ByVal countPtr, tilescount(0), 256 * LenB(tilescount(0))
    CountTiles = total
    
    
    On Error GoTo 0
    Exit Function

CountTiles_Error:
    HandleError Err, "frmMain.CountTiles"
End Function

Sub SetScrollbarValues(h As Long, V As Long, Optional Refresh As Boolean)
'Set the scrollbar values on the given coordinates
    
    
    
    Call UpdateScrollbars(Refresh)
    
    dontUpdateOnValueChange = Not Refresh
    
'    If Hscr.Enabled Then
'    If Hscr.Enabled Then
'
'    End If
    
        If h > hScr.Max Then
            hScr.value = hScr.Max
        ElseIf h < hScr.Min Then
            hScr.value = hScr.Min
        Else
            hScr.value = h
        End If
'    End If
    
'    If Vscr.Enabled Then
        If V > vScr.Max Then
            vScr.value = vScr.Max
        ElseIf V < vScr.Min Then
            vScr.value = vScr.Min
        Else
            vScr.value = V
        End If
'    End If
    
    dontUpdateOnValueChange = False
End Sub

Sub UpdateScrollbars(Optional Refresh As Boolean = False)
'put the scrollbars, but don't update the level
    
    dontUpdateOnValueChange = Not Refresh

    If 1024& * (currenttilew) - (picPreview.width \ currenttilew) * currenttilew <= 0 Then
        hScr.Max = 0
        hScr.value = 0
        hScr.Enabled = False
    Else
        hScr.Enabled = True
        hScr.Max = 1024& * currenttilew - (picPreview.width \ currenttilew) * currenttilew
    End If
    hScr.LargeChange = currenttilew * 10
    hScr.SmallChange = currenttilew


    If 1024& * (currenttilew) - (picPreview.height \ currenttilew) * currenttilew <= 0 Then
        vScr.Max = 0
        vScr.value = 0
        vScr.Enabled = False
    Else
        vScr.Enabled = True
        vScr.Max = 1024& * currenttilew - (picPreview.height \ currenttilew) * currenttilew
    End If
    vScr.LargeChange = currenttilew * 10
    vScr.SmallChange = currenttilew
    
    
    dontUpdateOnValueChange = False
End Sub





Function WorldToTile(world As Integer) As Integer
    WorldToTile = world \ TILEW
End Function

Function WorldToScreenX(worldX As Integer) As Integer
    WorldToScreenX = (worldX * currentzoom) - hScr.value
End Function

Function WorldToScreenY(worldY As Integer) As Integer
    WorldToScreenY = (worldY * currentzoom) - vScr.value
End Function

Function TileToWorld(tile As Integer) As Integer
    TileToWorld = tile * TILEW
End Function

Function TileToScreenX(tileX As Integer) As Integer
    TileToScreenX = tileX * currenttilew - hScr.value
End Function

Function TileToScreenY(tileY As Integer) As Integer
    TileToScreenY = tileY * currenttilew - vScr.value
End Function

Function ScreenToTileX(screenX As Integer) As Integer
    ScreenToTileX = (hScr.value + screenX) \ currenttilew
End Function

Function ScreenToTileY(screenY As Integer) As Integer
    ScreenToTileY = (vScr.value + screenY) \ currenttilew
End Function

Function GetCornerOfTileByScreenX(screenX As Integer) As Integer
    GetCornerOfTileByScreenX = ((hScr.value + screenX) \ currenttilew) * currenttilew - hScr.value
End Function

Function GetCornerOfTileByScreenY(screenY As Integer) As Integer
    GetCornerOfTileByScreenY = ((vScr.value + screenY) \ currenttilew) * currenttilew - vScr.value
End Function




Sub RedrawRegionsArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, updateBuffer As Boolean)
    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_Regions)
    
    
    Dim firstscreenX As Integer, firstscreenY As Integer, _
        curscreenX As Integer, curscreenY As Integer, OffsetX As Integer, OffsetY As Integer
        
    firstscreenX = WorldToScreenX(Left)
    firstscreenY = WorldToScreenY(Top)
    
    Call layer.EraseArea(firstscreenX, _
                            firstscreenY, _
                            firstscreenX + (Right - Left) * currentzoom, _
                            firstscreenY + (Bottom - Top) * currentzoom)


    Dim layerHdc As Long
    layerHdc = layer.hDC
    
    
    Dim firstTileX As Integer, firstTileY As Integer, _
        curtilex As Integer, curtiley As Integer
    firstTileX = WorldToTile(Left)
    firstTileY = WorldToTile(Top)
    
    OffsetX = (Left Mod TILEW) * currentzoom
    OffsetY = (Top Mod TILEH) * currentzoom
    
    Call Regions.DrawOnLayer(layerHdc, _
                            firstscreenX, firstscreenY, _
                            (Right - Left) * currentzoom, (Bottom - Top) * currentzoom, _
                            firstTileX, firstTileY, OffsetX, OffsetY)
    
'    Exit Sub
'
'
'    Dim i As Integer, j As Integer
'
''    If UsingPixels Then
'''        'just copy a part of the pixel map if we are using pixels
''
''    Else
'        curtiley = firstTileY
'        curscreenY = firstscreenY
'
'        For j = Top To Bottom Step TILEH
'            curtilex = firstTileX
'            curscreenX = firstscreenX
'
'
'            If curtiley <= 1023 Then
'                For i = Left To Right Step TILEW
'                    If curtilex <= 1023 Then
'
'                        Call Regions.DrawRegionOn(layerhDC, curtilex, curtiley, curscreenX, curscreenY)
'
'                    End If
'
'                    curtilex = curtilex + 1
'                    curscreenX = curscreenX + currenttilew
'                Next i
'
'            End If
'
'            curtiley = curtiley + 1
'            curscreenY = curscreenY + currenttilew
'        Next j
    
'    End If
    
    DrawRegions = ShowRegions And Regions.HaveVisibleRegions
                            
    If updateBuffer Then Call RedrawBufferArea(firstscreenX, _
                                                firstscreenY, _
                                                WorldToScreenX(Right), _
                                                WorldToScreenY(Bottom), True)
                                                
    Set layer = Nothing
End Sub

Private Sub RedrawLvz(firstLayer As LVZLayerEnum, lastLayer As LVZLayerEnum, ByRef drawOn As clsDisplayLayer, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer)
    Dim firstscreenX As Integer, firstscreenY As Integer, layer As Integer
    Dim area As RECT
    
    area.Left = Left
    area.Right = Right
    area.Top = Top
    area.Bottom = Bottom
    
    firstscreenX = WorldToScreenX(Left)
    firstscreenY = WorldToScreenY(Top)
    
    Call drawOn.EraseArea(firstscreenX, _
                            firstscreenY, _
                            firstscreenX + (Right - Left) * currentzoom, _
                            firstscreenY + (Bottom - Top) * currentzoom)
    
    For layer = firstLayer To lastLayer
    
        Call lvz.DrawLVZsInArea(drawOn.hDC, layer, area, firstscreenX, firstscreenY)
    
    Next
    
End Sub



Sub RedrawLvzUnderArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, updateBuffer As Boolean)
    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_LVZunder)
    
    Call RedrawLvz(lyr_BelowAll, lyr_AfterBackground, layer, Left, Top, Right, Bottom)
    
    DrawLVZ = ShowLVZ And lvz.HaveLVZ
    
    If updateBuffer Then Call RedrawBufferArea(WorldToScreenX(Left), _
                                                WorldToScreenY(Top), _
                                                WorldToScreenX(Right), _
                                                WorldToScreenY(Bottom), True)
    
    Set layer = Nothing
End Sub

Sub RedrawLvzOverArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, updateBuffer As Boolean)
    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_LVZover)
    
    Call RedrawLvz(lyr_AfterTiles, lyr_TopMost, layer, Left, Top, Right, Bottom)
    
    DrawLVZ = ShowLVZ And lvz.HaveLVZ
    
    If updateBuffer Then Call RedrawBufferArea(WorldToScreenX(Left), _
                                                WorldToScreenY(Top), _
                                                WorldToScreenX(Right), _
                                                WorldToScreenY(Bottom), True)
    Set layer = Nothing
End Sub

Sub RedrawSelectionArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, updateBuffer As Boolean)
    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_Selection)
    
    Dim layerHdc As Long
    layerHdc = layer.hDC
    
    layer.BackColor = IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor)
    
    
    Dim firstTileX As Integer, firstTileY As Integer, _
        curtilex As Integer, curtiley As Integer
    Dim firstscreenX As Integer, firstscreenY As Integer, _
        lastscreenX As Integer, lastscreenY As Integer, _
        curscreenX As Integer, curscreenY As Integer
    
    firstTileX = WorldToTile(Left)
    firstTileY = WorldToTile(Top)
    
    firstscreenX = TileToScreenX(firstTileX)
    firstscreenY = TileToScreenY(firstTileY)
    
    lastscreenX = WorldToScreenX(Right) '  firstscreenX + (Right - Left) * currentzoom
    lastscreenY = WorldToScreenY(Bottom) ' firstscreenY + (Bottom - Top) * currentzoom
    
    Call layer.EraseArea(firstscreenX, _
                            firstscreenY, _
                            lastscreenX, _
                            lastscreenY)
    
    'Stop now if there's no selection
    If Not sel.hasAlreadySelectedParts Then Exit Sub
    
    
    Dim tileid As Integer
    
    
    Dim UsingPixels As Boolean
    UsingPixels = magnifier.UsingPixels
    
    Dim objX As Integer, objY As Integer
    
    
    
    
    Dim i As Integer, j As Integer
    
    Dim selbounds As area
    
    If UsingPixels Then
'        'just copy a part of the pixel map if we are using pixels
        Dim bltwidth As Integer, bltheight As Integer
        bltwidth = intMinimum(((Right - Left) \ 16) + 1, 1024)
        bltheight = intMinimum(((Bottom - Top) \ 16) + 1, 1024)
        
        selbounds = sel.getBoundaries
        
        If bltwidth > 0 And bltheight > 0 Then
            If pastetype = p_trans Then
                TransparentBlt layerHdc, firstscreenX, firstscreenY, bltwidth, bltheight, pic1024selection.hDC, firstTileX, firstTileY, vbWhite
            Else
                BitBlt layerHdc, firstscreenX, firstscreenY, bltwidth, bltheight, pic1024selection.hDC, firstTileX, firstTileY, vbSrcCopy
            End If
            
            Call DrawRectangle(layerHdc, selbounds.Left - firstTileX + firstscreenX, _
                                    selbounds.Top - firstTileY + firstscreenY, _
                                    selbounds.Right - firstTileX + firstscreenX, _
                                    selbounds.Bottom - firstTileY + firstscreenY, _
                                    vbYellow)

        
        End If
        
        
        
        
        
    Else
        curtiley = firstTileY
'        curscreenY = firstscreenY
        
        selbounds = sel.getBoundaries
        
        Dim drawblack As Boolean
        drawblack = pastetype <> p_trans
        
        For curscreenY = firstscreenY To lastscreenY Step currenttilew
            curtilex = firstTileX
'            curscreenX = firstscreenX
            
            
            If curtiley <= selbounds.Bottom And curtiley >= selbounds.Top Then
                'For i = Left To Right Step TILEW
                For curscreenX = firstscreenX To lastscreenX Step currenttilew
                
                    If curtilex <= selbounds.Right And curtilex >= selbounds.Left Then
                        If sel.getIsInSelection(curtilex, curtiley) Then
                            tileid = sel.getSelTile(curtilex, curtiley)
                            
                            If tileid < 0 Then
                                'We might have a partial large object
                                'Find the top-left corner of the object
                                objX = ((tileid Mod 100) \ 10) + curtilex
                                objY = (curtiley + tileid Mod 10)
                                
                                'The only time we'll want to draw this object, is if it's not drawn yet
                                'For this, the top-left of the object must be outside our boundary
                                If objX < firstTileX Or objY < firstTileY Then
                                    'Ok, the object is partially out of our boundaries...
                                    
                                    'This tile has to be the first we encounter of the object
                                    If ((objX < firstTileX And curtilex = firstTileX) Or objX >= firstTileX) Then
                                        If ((objY < firstTileY And curtiley = firstTileY) Or objY >= firstTileY) Then
                                            'Draw partial object
                                            Call TileRender.DrawObject(tileid \ (-100), True, layer.hDC, curscreenX, curscreenY, curtilex - objX, curtiley - objY)
                                        End If
                                    End If
                                End If
    
                            Else
                                'Just a tile, draw it
                                Call TileRender.DrawTile(tileid, True, layer.hDC, curscreenX, curscreenY, False, drawblack)
                            End If
                            
                            Call sel.DrawCornerLinesOn(curtilex, curtiley, curscreenX, curscreenY, currenttilew, layer)
                        End If
                    End If
                    
                    curtilex = curtilex + 1
'                    curscreenX = curscreenX + currenttilew
                Next
                
            End If
            
            curtiley = curtiley + 1
'            curscreenY = curscreenY + currenttilew
        Next
    
    End If
    
    
    
    If updateBuffer Then Call RedrawBufferArea(firstscreenX, _
                                                firstscreenY, _
                                                WorldToScreenX(Right), _
                                                WorldToScreenY(Bottom), True)
    Set layer = Nothing
End Sub




Friend Sub RenderTiles(firsttileX As Integer, firsttileY As Integer, firstscreenX As Integer, firstscreenY As Integer, lastscreenX As Integer, lastscreenY As Integer, renderhDC As Long, usepixels As Boolean, usetilenr As Boolean)

    Dim curscreenX As Integer, curscreenY As Integer, curtilex As Integer, curtiley As Integer
    
    Dim objX As Integer, objY As Integer
    Dim i As Integer, j As Integer
    Dim tileid As Integer
    
    
    If usepixels Then
        Dim bltwidth As Integer, bltheight As Integer
        bltwidth = intMinimum(lastscreenX - firstscreenX + 1, 1024 - firsttileX)
        bltheight = intMinimum(lastscreenY - firstscreenY + 1, 1024 - firsttileY)
        
        If bltwidth > 0 And bltheight > 0 Then
'        'just copy a part of the pixel map if we are using pixels
'        'BitBlt piclevelhdc, 0, 0, piclevel.width, piclevel.height, pic1024.hdc, lbx, lby, vbSrcCopy
            Call cpic1024.transToDC(renderhDC, firstscreenX, firstscreenY, bltwidth, bltheight, firsttileX, firsttileY, vbBlack)
'        Call cpic1024.bltToDC(layer.hDC, 0, 0, layer.width, layer.height, firsttileX, firsttileY, vbSrcCopy)
        End If
    Else
        curtiley = firsttileY
'        curscreenY = firstscreenY
        
        If usingtilenr Then
            Call ChangeTextSize(renderhDC, 8)
            Call ChangeTextColor(renderhDC, vbWhite)
        End If
        
        For curscreenY = firstscreenY To lastscreenY Step currenttilew
            curtilex = firsttileX
'            curscreenX = firstscreenX
            
            If curtiley < MAPH Then
                For curscreenX = firstscreenX To lastscreenX Step currenttilew
                    If curtilex < MAPW Then
                        tileid = tile(curtilex, curtiley)
                        
                        If tileid < 0 Then
                            'We might have a partial large object
                            'Find the top-left corner of the object
                            objX = ((tileid Mod 100) \ 10) + curtilex
                            objY = (curtiley + tileid Mod 10)
                            
                            'The only time we'll want to draw this object, is if it's not drawn yet
                            'For this, the top-left of the object must be outside our boundary
                            If objX < firsttileX Or objY < firsttileY Then
                                'Ok, the object is partially out of our boundaries...
                                
                                'This tile has to be the first we encounter of the object
                                If ((objX < firsttileX And curtilex = firsttileX) Or objX >= firsttileX) Then
                                    If ((objY < firsttileY And curtiley = firsttileY) Or objY >= firsttileY) Then
                                        'Draw partial object
                                        Call TileRender.DrawObject(tileid \ (-100), False, renderhDC, curscreenX, curscreenY, curtilex - objX, curtiley - objY)
                                    End If
                                End If
                            End If

                        ElseIf tileid > 0 Then
                            Call TileRender.DrawTile(tileid, False, renderhDC, curscreenX, curscreenY, True, False)
                            
                            If usetilenr Then
                                Call PrintText(renderhDC, CStr(tileid), CInt(curscreenX), CInt(curscreenY))
                            End If
                        End If
                    End If
                    
                    curtilex = curtilex + 1
'                    curscreenX = curscreenX + currenttilew
                Next
                
            End If
            
            curtiley = curtiley + 1
'            curscreenY = curscreenY + currenttilew
        Next
    
    End If
    
End Sub





Private Sub RedrawTileArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, updateBuffer As Boolean)
    'Redraws tiles on the Tiles layer in the given area

    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_Tiles)
    
    layer.BackColor = unusedTilesetColor
    TileRender.TransparencyColor = unusedTilesetColor
    
    Dim firsttileX As Integer, firsttileY As Integer
    
    firsttileX = WorldToTile(Left)
    firsttileY = WorldToTile(Top)
    Call RenderTiles(firsttileX, firsttileY, TileToScreenX(firsttileX), TileToScreenY(firsttileY), WorldToScreenX(Right), WorldToScreenY(Bottom), layer.hDC, magnifier.usingpixels, usingtilenr)

    
    If updateBuffer Then Call RedrawBufferArea(TileToScreenY(firsttileY), WorldToScreenX(Right), _
                                                WorldToScreenX(Right), _
                                                WorldToScreenY(Bottom), True)
    Set layer = Nothing
'    Call layer.EraseArea(Left, Top, Right, Bottom)
End Sub





Sub RedrawGridArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, updateBuffer As Boolean)
'Draws the grid
'layer : MapLayer object to draw grid on
'left, top, right, bottom : World coordinates (pixels) of the area to redraw

    Dim layer As clsDisplayLayer, layerHdc As Long
    Set layer = MapLayers(DL_Tiles)
    
    layerHdc = layer.hDC
    
    If magnifier.UsingPixels Then
        layer.BackColor = vbBlack
    Else
        layer.BackColor = unusedTilesetColor
    End If
    
    Dim i As Integer, j As Integer
    
    'offset, in pixels, from the tile to the grid line
    Dim OffsetX As Integer, OffsetY As Integer
    OffsetX = Left Mod TILEW
    OffsetY = Top Mod TILEH
        
    If OffsetX = 0 Then OffsetX = TILEW
    If OffsetY = 0 Then OffsetY = TILEH
    
    'calculate the width and height of the lines
    'fullw and fullh are needed to erase the full background
    Dim w As Integer, h As Integer, fullw As Integer, fullh As Integer

    fullw = (Right - Left + 1) * currentzoom
    If Right > MAPWpx Then
        w = (MAPWpx - Left) * currentzoom
    Else
        w = fullw
    End If

    fullh = (Bottom - Top + 1) * currentzoom
    If Bottom > MAPHpx Then
        h = (MAPHpx - Top) * currentzoom
    Else
        h = fullh
    End If

    

    Dim curtileH As Integer, curtileV As Integer
    Dim firstscreenH As Integer, firstscreenV As Integer, firstscreenH_adj As Integer, firstscreenV_adj As Integer
    
    Dim notusingpixels As Boolean
    notusingpixels = Not magnifier.UsingPixels
    
    curtileH = WorldToTile(Left + TILEW - OffsetX)
    curtileV = WorldToTile(Top + TILEH - OffsetY)
    
    firstscreenH = WorldToScreenX(Left)
    firstscreenV = WorldToScreenY(Top)
    
'    Call layer.EraseArea2(firstscreenH, firstscreenV, firstscreenH + w, firstscreenV + h)
    Call layer.EraseArea(firstscreenH, firstscreenV, firstscreenH + fullw + 1, firstscreenV + fullh + 1)
    
    firstscreenH_adj = WorldToScreenX(Left + TILEW - OffsetX)
    
    If (TestMap.isRunning And usinggridTest) Or usinggrid Then
        'Using grid
        
        'draw vertical lines
        For j = firstscreenH_adj To firstscreenH_adj + w Step currenttilew
            
            'Define the color of the grid line
            If curtileH = 0 Then
                Call DrawLine(layerHdc, j, firstscreenV, j, firstscreenV + h, vbWhite)
                
            ElseIf curtileH = 1024 Then
                Call DrawLine(layerHdc, j - 1, firstscreenV, j - 1, firstscreenV + h, vbWhite)
                
'                Call layer.DrawVerticalLine(j, firstscreenV, firstscreenV + h, vbWhite)
                
            ElseIf curtileH = 512 Then
                Call DrawLine(layerHdc, j, firstscreenV, j, firstscreenV + h, gridcolor(4))
                
            ElseIf curtileH Mod Int(gridblocksY * gridsectionsY) = GridOffsetY Then
                Call DrawLine(layerHdc, j, firstscreenV, j, firstscreenV + h, gridcolor(5))
                
            ElseIf curtileH Mod gridblocksY = GridOffsetY Then
                If notusingpixels Then Call DrawLine(layerHdc, j, firstscreenV, j, firstscreenV + h, gridcolor(6))
                
            ElseIf notusingpixels Then
                Call DrawLine(layerHdc, j, firstscreenV, j, firstscreenV + h, gridcolor(7))
            End If
            

            
            curtileH = curtileH + 1
    
        Next
    
        firstscreenV_adj = WorldToScreenY(Top + TILEH - OffsetY)
    
        'draw the horizontal lines
        For i = firstscreenV_adj To firstscreenV_adj + h Step currenttilew
            'Define the color of the grid line
            If curtileV = 0 Then
                Call DrawLine(layerHdc, firstscreenH, i, firstscreenH + w, i, vbWhite)
                
            ElseIf curtileV = 1024 Then
                Call DrawLine(layerHdc, firstscreenH, i - 1, firstscreenH + w, i - 1, vbWhite)
                
            ElseIf curtileV = 512 Then
                Call DrawLine(layerHdc, firstscreenH, i, firstscreenH + w, i, gridcolor(0))
'                Call layer.DrawHorizontalLine(firstscreenH, firstscreenH + w, i, gridcolor(0))
                
            ElseIf curtileV Mod Int(gridblocksX * gridsectionsX) = GridOffsetX Then
                Call DrawLine(layerHdc, firstscreenH, i, firstscreenH + w, i, gridcolor(1))
'                Call layer.DrawHorizontalLine(firstscreenH, firstscreenH + w, i, gridcolor(1))
                
            ElseIf curtileV Mod gridblocksX = GridOffsetX Then
                If notusingpixels Then Call DrawLine(layerHdc, firstscreenH, i, firstscreenH + w, i, gridcolor(2))
                
            ElseIf notusingpixels Then
                Call DrawLine(layerHdc, firstscreenH, i, firstscreenH + w, i, gridcolor(3))
            End If

            
            curtileV = curtileV + 1
        Next
        
    Else
        'Not using grid; just draw the map edges
        
        'Draw left edge
        If curtileH = 0 Then Call DrawFilledRectangle(layerHdc, firstscreenH, firstscreenV, firstscreenH + 1, firstscreenV + h, vbWhite)
                
        'Draw top edge
        If curtileV = 0 Then Call DrawFilledRectangle(layerHdc, firstscreenH, firstscreenV, firstscreenH + w, firstscreenV + 1, vbWhite)
        
        'Draw right edge
        If curtileH + w \ currenttilew = MAPW Then Call DrawFilledRectangle(layerHdc, firstscreenH + w, firstscreenV, firstscreenH + w + 1, firstscreenV + h, vbWhite)
        
        'Draw bottom edge
        If curtileV + h \ currenttilew = MAPH Then Call DrawFilledRectangle(layerHdc, firstscreenH, firstscreenV + h, firstscreenH + w + 1, firstscreenV + h + 1, vbWhite)
        
    End If

    If updateBuffer Then Call RedrawBufferArea(WorldToScreenX(Left), _
                                                WorldToScreenY(Top), _
                                                WorldToScreenX(Right), _
                                                WorldToScreenY(Bottom), True)
    
    Set layer = Nothing
    
End Sub

Sub RedrawBuffer(do_updatePreview As Boolean)
    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_Buffer)
    
    Call layer.Cls
    
    If DrawLVZ Then Call MapLayers(DL_LVZunder).BitBltToLayerFull(layer, vbSrcCopy)
    
    If DrawRegions And curtool <> T_Region Then Call MapLayers(DL_Regions).AlphaTransToLayerFull(layer, vbBlack, regnopacity2)
    
    If sel.hasAlreadySelectedParts Then
        If pastetype = p_under Then
            Call MapLayers(DL_Selection).TransparentBltToLayerFull(layer, IIf(magnifier.UsingPixels, vbWhite, unusedTilesetColor))
            Call MapLayers(DL_Tiles).TransparentBltToLayerFull(layer, IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor))
        Else
            Call MapLayers(DL_Tiles).TransparentBltToLayerFull(layer, IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor))
            Call MapLayers(DL_Selection).TransparentBltToLayerFull(layer, IIf(magnifier.UsingPixels, IIf(pastetype = p_trans, vbBlack, vbWhite), unusedTilesetColor))
        End If
    Else
        Call MapLayers(DL_Tiles).TransparentBltToLayerFull(layer, IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor))
    End If
    
    If DrawLVZ Then Call MapLayers(DL_LVZover).TransparentBltToLayerFull(layer, vbBlack)
    
    If DrawRegions And curtool = T_Region Then Call MapLayers(DL_Regions).AlphaTransToLayerFull(layer, vbBlack, regnopacity1)
    
    If do_updatePreview Then Call UpdatePreview
    
    Set layer = Nothing
End Sub

Sub RedrawBufferArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, do_updatePreview As Boolean)
    
    
    Dim layer As clsDisplayLayer
    Set layer = MapLayers(DL_Buffer)
    
    If Right < 0 Or Left > layer.width Or Bottom < 0 Or Top > layer.height Then
        Set layer = Nothing
        Exit Sub
    End If
    
    Call layer.EraseArea(Left, Top, Right, Bottom)

    If DrawLVZ Then Call MapLayers(DL_LVZunder).BitBltToLayer(layer, Left, Top, Right, Bottom, vbSrcCopy)
    
    If DrawRegions Then If curtool <> T_Region Then Call MapLayers(DL_Regions).AlphaTransToLayer(layer, Left, Top, Right, Bottom, vbBlack, regnopacity2)
    
    If sel.hasAlreadySelectedParts Then
        If pastetype = p_under Then
            Call MapLayers(DL_Selection).TransparentBltToLayer(layer, Left, Top, Right, Bottom, IIf(magnifier.UsingPixels, vbWhite, unusedTilesetColor))
            Call MapLayers(DL_Tiles).TransparentBltToLayer(layer, Left, Top, Right, Bottom, IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor))
        Else
            Call MapLayers(DL_Tiles).TransparentBltToLayer(layer, Left, Top, Right, Bottom, IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor))
            
            Call MapLayers(DL_Selection).TransparentBltToLayer(layer, Left, Top, Right, Bottom, IIf(magnifier.UsingPixels, IIf(pastetype = p_trans, vbBlack, vbWhite), unusedTilesetColor))
        End If
    Else
        Call MapLayers(DL_Tiles).TransparentBltToLayer(layer, Left, Top, Right, Bottom, IIf(magnifier.UsingPixels, vbBlack, unusedTilesetColor))
    End If
    
    If DrawLVZ Then Call MapLayers(DL_LVZover).TransparentBltToLayer(layer, Left, Top, Right, Bottom, vbBlack)
    
    If DrawRegions Then If curtool = T_Region Then Call MapLayers(DL_Regions).AlphaTransToLayer(layer, Left, Top, Right, Bottom, vbBlack, regnopacity1)
    
    If do_updatePreview Then Call UpdatePreviewArea(Left, Top, Right, Bottom, True)
    
    Set layer = Nothing
End Sub

Sub RedrawRegions(updateBuffer As Boolean)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    Call RedrawRegionsArea(lbx, lby, hbx, hby, False)
    
    If updateBuffer Then Call RedrawBuffer(True)
End Sub

Sub RedrawLvzUnderLayer(updateBuffer As Boolean)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    Call RedrawLvzUnderArea(lbx, lby, hbx, hby, False)
    
    If updateBuffer Then Call RedrawBuffer(True)
End Sub

Sub RedrawLvzOverLayer(updateBuffer As Boolean)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    Call RedrawLvzOverArea(lbx, lby, hbx, hby, False)
    
    If updateBuffer Then Call RedrawBuffer(True)
End Sub

Sub RedrawSelection(updateBuffer As Boolean)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    Call RedrawSelectionArea(lbx, lby, hbx, hby, False)
    
    If updateBuffer Then Call RedrawBuffer(True)
End Sub

Sub RedrawTileLayer(updateBuffer As Boolean)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    Call RedrawTileArea(lbx, lby, hbx, hby, False)
    
    If updateBuffer Then Call RedrawBuffer(True)
End Sub

Sub RedrawGrid(updateBuffer As Boolean)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    Call RedrawGridArea(lbx, lby, hbx, hby, False)
    
    If updateBuffer Then Call RedrawBuffer(True)
End Sub

Sub UpdateLevelArea(ByVal lbx As Integer, ByVal lby As Integer, ByVal hbx As Integer, ByVal hby As Integer)
    'Screen coordinates
'    Dim lbxScreen As Integer, hbxScreen As Integer, lbyScreen As Integer, hbyScreen As Integer
    
    If lbx Mod TILEW <> 0 Then lbx = lbx - (lbx Mod TILEW)
    If hbx Mod TILEW <> 0 Then hbx = hbx + TILEW - (hbx Mod TILEW)
    If lby Mod TILEH <> 0 Then lby = lby - (lby Mod TILEW)
    If hby Mod TILEH <> 0 Then hby = hby + TILEH - (hby Mod TILEW)
    
    
'    lbxScreen = WorldToScreenX(lbx)
'    lbyScreen = WorldToScreenY(lby)
'    hbxScreen = WorldToScreenX(hbx)
'    hbyScreen = WorldToScreenY(hby)
    
    'Redraw each layer
    If DrawRegions Then Call RedrawRegionsArea(lbx, lby, hbx, hby, False)
    
    
    If DrawLVZ Then Call RedrawLvzUnderArea(lbx, lby, hbx, hby, False)
    

    Call RedrawGridArea(lbx, lby, hbx, hby, False)
    Call RedrawTileArea(lbx, lby, hbx, hby, False)
    
    If sel.hasAlreadySelectedParts Then Call RedrawSelectionArea(lbx, lby, hbx, hby, False)
    
    
    If DrawLVZ Then Call RedrawLvzOverArea(lbx, lby, hbx, hby, False)
        
    
    
    'Build the buffer by combining all these
'    Call RedrawBufferArea(MapLayers(DL_Buffer), lbxScreen, lbyScreen, hbxScreen, hbyScreen)
    
End Sub


Sub UpdateLevelTest(ByRef tick() As Long, Optional redrawusingprevious As Boolean = False, Optional do_updatePreview As Boolean = True)
'Updates the level
'CALCULATE BOUNDARIES OF PICLEVEL
    
'    redrawusingprevious = False
    
    'World coordinates (pixels)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
    'World coordinates (tiles)
    Dim lbxTile As Integer, hbxTile As Integer, lbyTile As Integer, hbyTile As Integer
    
    'Screen coordinates
    Dim lbxScreen As Integer, hbxScreen As Integer, lbyScreen As Integer, hbyScreen As Integer
    

    
    Dim i As Integer
'    Dim j As Integer
    
    Static oldlbx As Integer
    Static oldlby As Integer
    'when resized only h change
    Static oldhbx As Integer
    Static oldhby As Integer
    
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom) - 1
    hby = Int((vScr.value + picPreview.height) / currentzoom) - 1
    
    DrawLVZ = ShowLVZ And lvz.HaveLVZ
    DrawRegions = ShowRegions And Regions.HaveVisibleRegions
    
    If redrawusingprevious Then
        'Grab part of the buffer


        
        Dim dx As Integer, dy As Integer
        dx = lbx - oldlbx
        dy = lby - oldlby
        
        
        'Move the 'known' part of each layer
        For i = DL_Regions To DL_Buffer
            With MapLayers(i)
                Call .MoveBitmap(-dx * currentzoom, -dy * currentzoom)
            End With
        Next
        
        'Calculate missing parts

        
'            If offsetX = 0 And offsetY = 0 And oldhbx = hbx And oldhby = hby Then
'                Exit Sub
'            End If


        If dx Or dy Then
            If dx > hbx - lbx Or dy > hby - lby Then
                Call UpdateLevel(False, do_updatePreview)
                Exit Sub
            End If
            
            If dx >= 0 Then
                If dy >= 0 Then
            '        +---------------+---+
            '        |               | dx|
            '        |               |   |
            '        |               | 1 |
            '        |               |   |
            '        +---------------+   +
            '        | dy   2        |   |
            '        +---------------+---+
                    If dx Then Call UpdateLevelArea(hbx - dx + 1, lby, hbx, hby)
                    If dy Then Call UpdateLevelArea(lbx, hby - dy + 1, hbx - dx, hby)
                Else
            '        +---------------+---+
            '        | -dy  1            |
            '        +---------------+---+
            '        |               | dx|
            '        |               |   |
            '        |               | 2 |
            '        |               |   |
            '        +---------------+---+
                    If dy Then Call UpdateLevelArea(lbx, lby, hbx, lby - dy - 1)
                    If dx Then Call UpdateLevelArea(hbx - dx + 1, lby - dy, hbx, hby)
                End If
            Else
                If dy >= 0 Then
            '        +---+---------------+
            '        |-dx|               |
            '        |   |               |
            '        | 1 |               |
            '        |   |               |
            '        |   +---------------+
            '        |   | dy 2          |
            '        +---+---------------+
                    If dx Then Call UpdateLevelArea(lbx, lby, lbx - dx - 1, hby)
                    If dy Then Call UpdateLevelArea(lbx - dx, hby - dy + 1, hbx, hby)
            
                Else
            '        +---+---------------+
            '        |   |-dy   2        |
            '        |   +---------------+
            '        |-dx|               |
            '        |   |               |
            '        | 1 |               |
            '        |   |               |
            '        +---+---------------+
                    If dx Then Call UpdateLevelArea(lbx, lby, lbx - dx - 1, hby)
                    If dy Then Call UpdateLevelArea(lbx - dx, lby, hbx, lby - dy - 1)
                
                End If
            End If
        End If
'            Dim li As Integer
'            Dim lj As Integer
'            Dim ui As Integer
'            Dim uj As Integer
'
'            If offsetX > 0 Then
'                li = lbx
'                ui = lbx + offsetX
'            Else
'                li = hbx + offsetX
'                ui = hbx
'            End If
'
'            If offsetY > 0 Then
'                lj = lby
'                uj = lby + offsetY
'            Else
'                lj = hby + offsetY
'                uj = hby
'            End If
        Call RedrawBuffer(False)
        
    Else
        'Calculate boundaries (in pixel)
        lbxTile = hScr.value \ currenttilew
        lbyTile = vScr.value \ currenttilew
        hbxTile = (hScr.value + picPreview.width) \ currenttilew
        hbyTile = (vScr.value + picPreview.height) \ currenttilew

        lbxScreen = 0
        lbyScreen = 0
        hbxScreen = picPreview.width
        hbyScreen = picPreview.height
        
        'Redraw each layer
        tick(DL_Regions) = GetTickCount
        
        If DrawRegions Then Call RedrawRegionsArea(lbx, lby, hbx, hby, False)
        
        tick(DL_Regions) = GetTickCount - tick(DL_Regions)
        
        tick(DL_LVZunder) = GetTickCount
        
        If DrawLVZ Then Call RedrawLvzUnderArea(lbx, lby, hbx, hby, False)
        
        tick(DL_LVZunder) = GetTickCount - tick(DL_LVZunder)
        
        tick(DL_Tiles) = GetTickCount
        
        Call RedrawGridArea(lbx, lby, hbx, hby, False)
        Call RedrawTileArea(lbx, lby, hbx, hby, False)
        
        tick(DL_Tiles) = GetTickCount - tick(DL_Tiles)
        
        tick(DL_Selection) = GetTickCount
        
        If sel.hasAlreadySelectedParts Then Call RedrawSelectionArea(lbx, lby, hbx, hby, False)
        
        tick(DL_Selection) = GetTickCount - tick(DL_Selection)
        
        tick(DL_LVZover) = GetTickCount
        
        If DrawLVZ Then Call RedrawLvzOverArea(lbx, lby, hbx, hby, False)
        
        tick(DL_LVZover) = GetTickCount - tick(DL_LVZover)
        
        'Build the buffer by combining all these
        tick(DL_Buffer) = GetTickCount
        
        Call RedrawBuffer(False)
        
        tick(DL_Buffer) = GetTickCount - tick(DL_Buffer)
    End If
    
    tick(DL_Buffer + 1) = GetTickCount
    
    If do_updatePreview Then Call UpdatePreview
    
    tick(DL_Buffer + 1) = GetTickCount - tick(DL_Buffer + 1)
    
    oldlbx = lbx
    oldlby = lby
    oldhbx = hbx
    oldhby = hby

End Sub



Sub UpdateLevel(Optional redrawusingprevious As Boolean = False, Optional do_updatePreview As Boolean = True)
'Updates the level
'CALCULATE BOUNDARIES OF PICLEVEL
    
'    redrawusingprevious = False
    
    'World coordinates (pixels)
    Dim lbx As Integer, hbx As Integer, lby As Integer, hby As Integer
    
'    'World coordinates (tiles)
'    Dim lbxTile As Integer, hbxTile As Integer, lbyTile As Integer, hbyTile As Integer
'
'    'Screen coordinates
'    Dim lbxScreen As Integer, hbxScreen As Integer, lbyScreen As Integer, hbyScreen As Integer
    

    
    Dim i As Integer
'    Dim j As Integer
    
    Static oldlbx As Integer
    Static oldlby As Integer
    'when resized only h change
    Static oldhbx As Integer
    Static oldhby As Integer
    
    
    lbx = hScr.value / currentzoom
    lby = vScr.value / currentzoom
    hbx = Int((hScr.value + picPreview.width) / currentzoom)
    hby = Int((vScr.value + picPreview.height) / currentzoom)
    

    
    DrawLVZ = ShowLVZ And lvz.HaveLVZ
    DrawRegions = ShowRegions And Regions.HaveVisibleRegions
    
    If redrawusingprevious Then
        'Grab part of the buffer


        
        Dim dx As Integer, dy As Integer
        dx = lbx - oldlbx
        dy = lby - oldlby
        
        
        'Move the 'known' part of each layer
        For i = DL_Regions To DL_Buffer
            With MapLayers(i)
                Call .MoveBitmap(-dx * currentzoom, -dy * currentzoom)
            End With
        Next
        
        'Calculate missing parts

        
'            If offsetX = 0 And offsetY = 0 And oldhbx = hbx And oldhby = hby Then
'                Exit Sub
'            End If


        If dx Or dy Then
            If dx > hbx - lbx Or dy > hby - lby Then
                Call UpdateLevel(False, do_updatePreview)
                Exit Sub
            End If
            
            If dx >= 0 Then
                If dy >= 0 Then
            '        +---------------+---+
            '        |               | dx|
            '        |               |   |
            '        |               | 1 |
            '        |               |   |
            '        +---------------+   +
            '        | dy   2        |   |
            '        +---------------+---+

                    If dx Then Call UpdateLevelArea(hbx - dx - 1, lby, hbx, hby)
                    If dy Then Call UpdateLevelArea(lbx, hby - dy - 1, hbx - dx, hby)
                Else
            '        +---------------+---+
            '        | -dy  1            |
            '        +---------------+---+
            '        |               | dx|
            '        |               |   |
            '        |               | 2 |
            '        |               |   |
            '        +---------------+---+
                    If dy Then Call UpdateLevelArea(lbx, lby, hbx, lby - dy - 1)
                    If dx Then Call UpdateLevelArea(hbx - dx + 1, lby - dy, hbx, hby)
                End If
            Else
                If dy >= 0 Then
            '        +---+---------------+
            '        |-dx|               |
            '        |   |               |
            '        | 1 |               |
            '        |   |               |
            '        |   +---------------+
            '        |   | dy 2          |
            '        +---+---------------+
                    If dx Then Call UpdateLevelArea(lbx, lby, lbx - dx - 1, hby)
                    If dy Then Call UpdateLevelArea(lbx - dx, hby - dy + 1, hbx, hby)
            
                Else
            '        +---+---------------+
            '        |   |-dy   2        |
            '        |   +---------------+
            '        |-dx|               |
            '        |   |               |
            '        | 1 |               |
            '        |   |               |
            '        +---+---------------+
                    If dx Then Call UpdateLevelArea(lbx, lby, lbx - dx - 1, hby)
                    If dy Then Call UpdateLevelArea(lbx - dx, lby, hbx, lby - dy - 1)
                
                End If
            End If
        ElseIf hby > oldhby Or hbx > oldhbx Then
            'It was resized
            '        +---------------+---+
            '        |               | dx|
            '        |               |   |
            '        |               | 1 |
            '        |               |   |
            '        +---------------+   +
            '        | dy   2        |   |
            '        +---------------+---+
            dx = hbx - oldhbx
            dy = hby - oldhby
            
            If dx Then Call UpdateLevelArea(hbx - dx + 1, lby, hbx, hby)
            If dy Then Call UpdateLevelArea(lbx, hby - dy + 1, hbx - dx, hby)
        End If
'            Dim li As Integer
'            Dim lj As Integer
'            Dim ui As Integer
'            Dim uj As Integer
'
'            If offsetX > 0 Then
'                li = lbx
'                ui = lbx + offsetX
'            Else
'                li = hbx + offsetX
'                ui = hbx
'            End If
'
'            If offsetY > 0 Then
'                lj = lby
'                uj = lby + offsetY
'            Else
'                lj = hby + offsetY
'                uj = hby
'            End If
        Call RedrawBuffer(False)
        
    Else
        'Calculate boundaries (in pixel)
'        lbxTile = hScr.value \ currenttilew
'        lbyTile = vScr.value \ currenttilew
'        hbxTile = (hScr.value + picPreview.width) \ currenttilew
'        hbyTile = (vScr.value + picPreview.height) \ currenttilew
'
'        lbxScreen = 0
'        lbyScreen = 0
'        hbxScreen = picPreview.width
'        hbyScreen = picPreview.height
        
        'Redraw each layer
        
        If DrawRegions Then Call RedrawRegionsArea(lbx, lby, hbx, hby, False)
        
        If DrawLVZ Then Call RedrawLvzUnderArea(lbx, lby, hbx, hby, False)
        
        Call RedrawGridArea(lbx, lby, hbx, hby, False)
        Call RedrawTileArea(lbx, lby, hbx, hby, False)
        
        If sel.hasAlreadySelectedParts Then Call RedrawSelectionArea(lbx, lby, hbx, hby, False)
        
        If DrawLVZ Then Call RedrawLvzOverArea(lbx, lby, hbx, hby, False)
        
        'Build the buffer by combining all these
        Call RedrawBuffer(False)
        
    End If

    If do_updatePreview Then Call UpdatePreview

    oldlbx = lbx
    oldlby = lby
    oldhbx = hbx
    oldhby = hby

End Sub



Sub UpdateLevelObject(X As Integer, Y As Integer, Optional Refresh As Boolean = True, Optional Draw As Boolean = True)
    Call cpic1024.setPixelLong(X, Y, TilePixelColor(tile(X, Y)))
    
    Dim tilenr As Integer, objsize As Integer
    tilenr = tile(X, Y)
    
    objsize = GetMaxSizeOfObject(tilenr)
    
    Dim color As Long
    color = TilePixelColor(tilenr)
    
    Dim i As Integer, j As Integer
    
    For i = X To X + objsize
        For j = Y To Y + objsize
            Call cpic1024.setPixelLong(i, j, color)
        Next
    Next
            
    If Draw And Not magnifier.UsingPixels Then
        Dim screenLeft As Integer, screenTop As Integer, screenRight As Integer, screenBottom As Integer
        
        screenLeft = TileToScreenX(X)
        screenTop = TileToScreenY(Y)

        Call TileRender.DrawObject(tilenr, False, MapLayers(DL_Tiles).hDC, screenLeft, screenTop, 0, 0)
        
'        Call RedrawTileArea(X * TILEW, Y * TILEH, (X + 1) * TILEW, (Y + 1) * TILEH, Refresh)
        

'
        screenRight = screenLeft + (objsize + 1) * currenttilew
        screenBottom = screenTop + (objsize + 1) * currenttilew
'
        Call RedrawBufferArea(screenLeft, screenTop, screenRight, screenBottom, Refresh)
'
'
'        If Refresh Then
'            Call UpdatePreviewArea(screenLeft, screenTop, screenRight, screenBottom, True)
'        End If
    End If
End Sub

Sub UpdateLevelTile(X As Integer, Y As Integer, Optional Refresh As Boolean = True, Optional Draw As Boolean = True)
'Update a specific tile
'    On Error GoTo UpdateLevelTile_Error

    'Call setPixel(pic1024.hdc, x, y, TilePixelColor(tile(x, y)))
    Call cpic1024.setPixelLong(X, Y, TilePixelColor(tile(X, Y)))
    

    If Draw And Not magnifier.UsingPixels Then
        Dim screenLeft As Integer, screenTop As Integer, screenRight As Integer, screenBottom As Integer
        
        screenLeft = TileToScreenX(X)
        screenTop = TileToScreenY(Y)
        
        Call TileRender.DrawTile(tile(X, Y), False, MapLayers(DL_Tiles).hDC, screenLeft, screenTop, True, True)
        
'        Call RedrawTileArea(MapLayers(DL_Tiles), X * TILEW, Y * TILEH, (X + 1) * TILEW - 1, (Y + 1) * TILEH - 1)
        
        

        screenRight = screenLeft + currenttilew - 1
        screenBottom = screenTop + currenttilew - 1
        
        
'        If Refresh Then
        Call RedrawBufferArea(screenLeft, screenTop, screenRight, screenBottom, Refresh)
'            Call UpdatePreviewArea(screenLeft, screenTop, screenRight, screenBottom, True)
'        End If
    End If
    
End Sub




Sub UpdatePreviewArea(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, doUpdateRadar As Boolean)
    If TileText.isActive Then Call TileText.UpdateTextPreview(True, False)
    
    BitBlt picPreview.hDC, Left, Top, Right - Left + 1, Bottom - Top + 1, MapLayers(DL_Buffer).hDC, Left, Top, vbSrcCopy
    
    picPreview.Refresh
    
    If doUpdateRadar Then Call UpdateRadar
End Sub


Private Sub UpdateRadar()
    With frmGeneral.picradar
        
        If .width > 0 And .height > 0 Then
        
        
            '///////////////// RADAR UPDATE ////////////////
    
            'calculates the lbx and lby of the preview
            Dim maplbx As Integer, maplby As Integer, mapsizex As Integer, mapsizey As Integer
        
            maplbx = hScr.value \ currenttilew
            maplby = vScr.value \ currenttilew
            mapsizex = picPreview.width \ currenttilew
            mapsizey = picPreview.height \ currenttilew


            If UsingRadarFullMap Then
                SetStretchBltMode .hDC, HALFTONE
    
                'StretchBlt frmGeneral.picradar.hdc, 0, 0, frmGeneral.picradar.width, frmGeneral.picradar.height, pic1024.hdc, 0, 0, 1024, 1024, vbSrcCopy
                Call cpic1024.stretchToDC(.hDC, 0, 0, .width, .height, 0, 0, 1024, 1024, vbSrcCopy)
                
'                TransparentBlt
                
                Dim displayRatio As Double
                displayRatio = .width / 1024
                Call DrawRectangle(frmGeneral.picradar.hDC, Int(displayRatio * maplbx) - 1, Int(displayRatio * maplby) - 1, Int(displayRatio * (maplbx + mapsizex)) + 1, Int((mapsizey + maplby) * displayRatio) + 1, vbRed)
            Else
                Dim centralx As Integer, centraly As Integer
    
    
                'clear the picradar
                frmGeneral.picradar.Cls
    
                'calculate the central tile x and tile y
                centralx = (hScr.value \ currenttilew) + (picPreview.width / currenttilew) \ 2
                centraly = (vScr.value \ currenttilew) + (picPreview.height / currenttilew) \ 2
    
                Dim lbx As Integer, lby As Integer, bltwidth As Integer, bltheight As Integer
                'calculates the lbx and lby of the radar
    
                lbx = centralx - (.width \ 2) - 2
                lby = centraly - (.height \ 2) - 2
    
                
    
                'if we click on the radar, we still need to know our position
                'of the preview ... actually come to think of it, i could just
                'recalculate it from preview, but this is more .. err.. certain
                'that it will work
                
                
                radar_left = lbx
                radar_top = lby
                
                If lbx < 0 Then lbx = 0
                If lby < 0 Then lby = 0
                
                bltwidth = intMinimum(1024, radar_left + .width) - lbx
                bltheight = intMinimum(1024, radar_top + .height) - lby
                
                
                'draw the correct portion of the pixel map to the radar
                'Call BitBlt(frmGeneral.picradar.hdc, 0, 0, frmGeneral.picradar.width - 1, frmGeneral.picradar.height - 1, pic1024.hdc, Int(lbx) - 1, Int(lby) - 1, vbSrcCopy)
                
                Call cpic1024.bltToDC(.hDC, lbx - radar_left, lby - radar_top, bltwidth, bltheight, lbx, lby, vbSrcCopy)
    
                If sel.hasAlreadySelectedParts Then Call TransparentBlt(.hDC, lbx - radar_left, lby - radar_top, bltwidth, bltheight, pic1024selection.hDC, lbx, lby, vbWhite)
    
    
                'draw the center lines with dotted blue lines
                If radar_top > 512 - .height And radar_top < 512 Then
                    .ForeColor = vbBlue
                    .DrawStyle = 2
                    frmGeneral.picradar.Line (0, 512 - radar_top)-(.width, 512 - radar_top)
                    .DrawStyle = 0
                End If
                If radar_left > 512 - frmGeneral.picradar.width And radar_left < 512 Then
                    .ForeColor = vbBlue
                    .DrawStyle = 2
                    frmGeneral.picradar.Line (512 - radar_left, 0)-(512 - radar_left, frmGeneral.picradar.height)
                    .DrawStyle = 0
                End If
                
                ''fill the boundaries of the level on the map with grey lines
                If radar_left < 0 Then
                    Call DrawFilledRectangle(.hDC, 0, 0, Abs(radar_left), .height - 1, RADAR_OUTSIDE_COLOR)
                End If
                
                If radar_left + .width > MAPW Then
                    Call DrawFilledRectangle(.hDC, MAPW - radar_left + 1, 0, .width - 1, .height - 1, RADAR_OUTSIDE_COLOR)
                End If
                
                If radar_top < 0 Then
                    Call DrawFilledRectangle(.hDC, 0, 0, .width, Abs(radar_top), RADAR_OUTSIDE_COLOR)
                End If
                
                If radar_top + .height > MAPH Then
                    Call DrawFilledRectangle(.hDC, 0, MAPH - radar_top + 1, .width, .height - 1, RADAR_OUTSIDE_COLOR)
                End If
'                If lby + frmGeneral.picradar.height > 1024 Then
'                    frmGeneral.picradar.ForeColor = RGB(80, 80, 80)
'                    For i = 1024 - lby + 1 To frmGeneral.picradar.height
'                        frmGeneral.picradar.Line (0, i)-(frmGeneral.picradar.width, i)
'                    Next i
'                End If
    
    
                Call DrawRectangle(frmGeneral.picradar.hDC, _
                                   maplbx - radar_left, _
                                   maplby - radar_top, _
                                   maplbx - radar_left + mapsizex + 1, _
                                   maplby - radar_top + mapsizey + 1, vbRed)
    
    
            End If
    
            'refresh the radar
            .Refresh

        End If
    End With
End Sub

Sub UpdatePreview(Optional Refresh As Boolean = True, Optional radar As Boolean = True)
'Updates the preview and radar
    
    BitBlt picPreview.hDC, 0, 0, picPreview.width, picPreview.height, MapLayers(DL_Buffer).hDC, 0, 0, vbSrcCopy
    
    If TileText.isActive Then Call TileText.UpdateTextPreview(True, False)
    
    Call lvz.ShowSelection(True)
    
'          Call PaintPreview
    If Refresh Then picPreview.Refresh

    If radar Then Call UpdateRadar
    

    
End Sub










Sub UseShortcutTool(KeyDown As Boolean, KeyCode As Integer, Shift As Integer)
'Uses hotkeys to swap to different tools and such
    On Error GoTo UseShortcutTool_Error

    If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
        'use numpad for the bookmarks
        KeyCode = KeyCode + (vbKey0 - vbKeyNumpad0)
    End If

    If KeyDown And Not Shift = 2 And Not Shift = 1 Then
        Select Case KeyCode
        Case vbKeyEscape
            'stop selection when pressing esc
            If curtool = T_selection And sel.selstate = Append Then
                Call sel.StopSelecting
            End If

        Case vbKeyDelete
            'deletes the selection
            If curtool = T_selection And sel.selstate = Append Then
                Call ClearSelection
            ElseIf curtool = T_lvz Then
                If lvz.HasSelection Then
                  Call lvz.DeleteSelectedObject
'                  Call RedrawLvzUnderLayer(False)
'                  Call RedrawLvzOverLayer(True)
                End If
            End If

        Case vbKeyH
            frmGeneral.SetCurrentTool (T_hand)
        Case vbKeyP
            frmGeneral.SetCurrentTool (T_pencil)
        Case vbKeyE
            frmGeneral.SetCurrentTool (T_Eraser)
        Case vbKeyS
            frmGeneral.SetCurrentTool (T_selection)
        Case vbKeyD
            'frmGeneral.SetCurrentTool (T_dropper)
            tempdropping = True
            If curtool <> T_dropper Then
                toolbeforetempdropping = curtool
            End If
            Call frmGeneral.SetCurrentTool(T_dropper)

        Case vbKeyZ
            frmGeneral.SetCurrentTool (T_magnifier)
        Case vbKeyF
            frmGeneral.SetCurrentTool (T_bucket)
        Case vbKeyL
            frmGeneral.SetCurrentTool (T_line)
        Case vbKeyR
            frmGeneral.SetCurrentTool (T_rectangle)
        Case vbKeyO
            frmGeneral.SetCurrentTool (T_ellipse)
        Case vbKeyI
            frmGeneral.SetCurrentTool (t_spline)
        Case vbKeyA
            frmGeneral.SetCurrentTool (T_airbrush)
        Case vbKeyB
            frmGeneral.SetCurrentTool (T_replacebrush)
        Case vbKeyW
            frmGeneral.SetCurrentTool (T_magicwand)
        Case vbKeyG
            frmGeneral.ToggleGrid
        Case vbKeyT
            frmGeneral.ToggleTileNr
        Case vbKeyLeft
            'move the map 10 tiles to the left
            If hScr.value - (10 * TILEW) < 0 Then
                hScr.value = 0
            Else
                hScr.value = hScr.value - (10 * TILEW)
            End If

        Case vbKeyRight
            'move the map 10 tiles to the right
            If hScr.value + (10 * TILEW) > hScr.Max Then
                hScr.value = hScr.Max
            Else
                hScr.value = hScr.value + (10 * TILEW)
            End If

        Case vbKeyUp
            'move the map 10 tiles up
            If vScr.value - (10 * TILEH) < 0 Then
                vScr.value = 0
            Else
                vScr.value = vScr.value - (10 * TILEH)
            End If

        Case vbKeyDown
            'move the map 10 tiles down
            If vScr.value + (10 * TILEH) > vScr.Max Then
                vScr.value = vScr.Max
            Else
                vScr.value = vScr.value + (10 * TILEH)
            End If

        Case vbKey0 To vbKey9
            Call GotoBookMark(KeyCode - vbKey0)
        End Select

    ElseIf KeyDown And Shift = 1 And KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
        Call SetBookMark(KeyCode - vbKey0, (hScr.value + picPreview.width / 2) / (currenttilew), (vScr.value + picPreview.height / 2) / (currenttilew))
    ElseIf KeyDown And Shift = 2 And (curtool <> T_dropper And curtool <> T_selection And curtool <> T_magicwand And curtool <> T_Region And curtool <> T_freehandselection And curtool <> T_lvz) And SharedVar.MouseDown = 0 Then
        'if we press ctrl and we use pencil or bucket, use the dropper
        tempdropping = True
        toolbeforetempdropping = curtool
        Call frmGeneral.SetCurrentTool(T_dropper)
    Else
        'not pressing ctrl anymore, use the tool before dropping
        If tempdropping And Not IsControl(Shift) Then
            tempdropping = False
            Call frmGeneral.SetCurrentTool(toolbeforetempdropping)
        End If
    End If

    If Not KeyDown And (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then
        Call UpdateLevel(False)
    End If

    On Error GoTo 0
    Exit Sub

UseShortcutTool_Error:
    HandleError Err, "frmMain.UseShortcutTool"
End Sub

Sub SetBookMark(Index As Integer, Optional X As Integer = -1, Optional Y As Integer = -1)
    If X = -1 Then X = (hScr.value + picPreview.width / 2) / (currenttilew)
    If Y = -1 Then Y = (vScr.value + picPreview.height / 2) / (currenttilew)
    bookmark(Index) = CLng(Y) * 1024 + X

    Call frmGeneral.UpdateToolBarButtons
End Sub

Function BookmarkInfo(Index As Integer) As String
    BookmarkInfo = Index & ": Goto " & bookmark(Index) Mod 1024 & " , " & bookmark(Index) \ 1024
End Function

Function GetBookMarkX(Index As Integer) As Integer
    GetBookMarkX = bookmark(Index) Mod 1024
End Function

Function GetBookMarkY(Index As Integer) As Integer
    GetBookMarkY = bookmark(Index) \ 1024
End Function

Function GetBookMarkDefaultX(Index As Integer) As Integer
    If Index = 0 Or Index = 2 Or Index = 5 Or Index = 8 Then
        GetBookMarkDefaultX = 512
    ElseIf Index = 1 Or Index = 4 Or Index = 7 Then
        GetBookMarkDefaultX = 0
    Else
        GetBookMarkDefaultX = 1023
    End If
    
End Function

Function GetBookMarkDefaultY(Index As Integer) As Integer
    If Index = 0 Or Index = 4 Or Index = 5 Or Index = 6 Then
        GetBookMarkDefaultY = 512
    ElseIf Index = 7 Or Index = 8 Or Index = 9 Then
        GetBookMarkDefaultY = 0
    Else
        GetBookMarkDefaultY = 1023
    End If
End Function

Sub GotoBookMark(Index As Integer)
    Call SetFocusAt(GetBookMarkX(Index), GetBookMarkY(Index), picPreview.width \ 2, picPreview.height \ 2, True)
End Sub
Private Sub CheckShiftCtrl(Shift As Integer, curtilex As Integer, curtiley As Integer)
'Checks if we pressed shift or ctrl
    Dim diagdist As Integer

    tempx = curtilex
    tempy = curtiley
    If IsShift(Shift) Then
        'shift is pressed
        usingshift = True
        Select Case curtool
        Case T_line, t_spline, T_customshape
            ' calculate the distance from the current vector to the diagonal
            diagdist = 0.25 * (((curtilex - fromtilex) ^ 2) + ((curtiley - fromtiley) ^ 2)) ^ 0.5
            
            If (Abs(curtiley - fromtiley) - Abs(curtilex - fromtilex)) < diagdist And (Abs(curtiley - fromtiley) - Abs(curtilex - fromtilex)) > -diagdist Then
                If ((curtiley - fromtiley) < 0 And (curtilex - fromtilex) > 0) Or ((curtiley - fromtiley) > 0 And (curtilex - fromtilex) < 0) Then
                    tempy = fromtiley - (curtilex - fromtilex)    ' / Diagonals
                Else
                    tempy = fromtiley + (curtilex - fromtilex)    ' \ Diagonals
                End If

            Else
                If Abs(curtiley - fromtiley) > Abs(curtilex - fromtilex) Then    'Straight lines
                    tempx = fromtilex
                Else
                    tempy = fromtiley
                End If
            End If

        Case T_rectangle, T_ellipse, T_filledellipse, T_filledrectangle, T_selection
            If ((curtiley - fromtiley) < 0 And (curtilex - fromtilex) > 0) Or ((curtiley - fromtiley) > 0 And (curtilex - fromtilex) < 0) Then
                If Abs(curtilex - fromtilex) > Abs(curtiley - fromtiley) Then
                    tempy = fromtiley - (curtilex - fromtilex)
                Else
                    tempx = fromtilex - (curtiley - fromtiley)
                End If
            Else
                If Abs(curtilex - fromtilex) > Abs(curtiley - fromtiley) Then
                    tempy = fromtiley + (curtilex - fromtilex)
                Else
                    tempx = fromtilex + (curtiley - fromtiley)
                End If
            End If
        End Select
    Else
        usingshift = False
    End If

    usingctrl = (curtool <> T_customshape And IsControl(Shift))


End Sub

Private Sub DragMap(Button As Integer, ByVal X As Single, ByVal Y As Single, Shift As Integer)
    Dim Speed As Integer
    Dim Refresh As Boolean
    Dim Factor As Integer
    Factor = 5    'each Factor pixels increase speed by 1

    'Speed up the process if holding Alt key
    On Error GoTo DragMap_Error

    dontUpdateOnValueChange = True
    If X < 0 Then
        Speed = Abs(X) / DRAG_FACTOR_LEFT

        'move the map X to the left
        If hScr.value - Speed * (currenttilew) < 0 Then
            If hScr.value <> 0 Then
                hScr.value = 0
                Refresh = True
            End If
        Else
            hScr.value = hScr.value - Speed * (currenttilew)
            Refresh = True
        End If
        X = 0
    ElseIf X > picPreview.width Then
        Speed = (X - picPreview.width) / DRAG_FACTOR_RIGHT

        'move the map X tile to the right
        If hScr.value + Speed * (currenttilew) > hScr.Max Then
            If hScr.value <> hScr.Max Then
                hScr.value = hScr.Max
                Refresh = True
            End If
        Else
            hScr.value = hScr.value + Speed * (currenttilew)
            Refresh = True
        End If
        X = picPreview.width
    End If

    If Y < 0 Then
        Speed = Abs(Y) / DRAG_FACTOR_TOP

        'move the map X tile up
        If vScr.value - Speed * (currenttilew) < 0 Then
            If vScr.value <> 0 Then
                vScr.value = 0
                Refresh = True
            End If
        Else
            vScr.value = vScr.value - Speed * (currenttilew)
            Refresh = True
        End If
        
        Y = 0
    ElseIf Y > picPreview.height Then
        Speed = (Y - picPreview.height) / DRAG_FACTOR_BOTTOM

        'move the map X tile down
        If vScr.value + Speed * (currenttilew) > vScr.Max Then
            If vScr.value <> 0 Then
                vScr.value = vScr.Max
                Refresh = True
            End If
        Else
            vScr.value = vScr.value + Speed * (currenttilew)
            Refresh = True
        End If
        
        Y = picPreview.height
    End If
    dontUpdateOnValueChange = False

    If Refresh Then
'        Call picPreview_MouseMove(Button, Shift, x, y)   <--wtf? Infinite recursive madness
        Call UpdateLevel(False)
    End If
    
    On Error GoTo 0
    Exit Sub

DragMap_Error:
    HandleError Err, "frmMain.DragMap"
End Sub


Sub InitfrmWalltiles()
    Call frmCreateWallTile.setParent(Me)
    Call StoreTempWallTiles
End Sub

Sub StoreTempWallTiles()
    Dim i As Integer
    Dim a As Integer

    For i = 0 To 7
        For a = 0 To 15
            Call frmCreateWallTile.tempwalltiles.setWallTile(i, a, walltiles.getWallTile(i, a))
        Next
    Next
    Call frmCreateWallTile.Init(walltiles.curwall)
End Sub


Sub CopyWallTiles(newcurwall As Integer)
    Dim i As Integer
    Dim a As Integer

    For i = 0 To 7
        For a = 0 To 15
            Call walltiles.setWallTile(i, a, frmCreateWallTile.tempwalltiles.getWallTile(i, a))
        Next
    Next

    Call walltiles.SetCurwall(newcurwall)

    Call walltiles.SetTileIsWalltile

    Call tileset.SelectWalltiles(vbLeftButton, newcurwall, False)
    
    Call frmGeneral.UpdateToolBarButtons
End Sub


Private Sub initCoordLabels(Optional Button As Integer)

    If curtool = T_hand Or curtool = T_bucket Or curtool = T_magnifier Then
        frmGeneral.lblFrom = "X= " & fromtilex & " - Y= " & fromtiley

    ElseIf curtool = T_airbrush Or curtool = T_dropper Or curtool = T_Eraser Or curtool = T_pencil Then
        frmGeneral.lblToA.Caption = "Distance:"
        frmGeneral.lblTo.Caption = "X= " & 0 & " - Y= " & 0
        frmGeneral.lblFrom = "X= " & fromtilex & " - Y= " & fromtiley

    ElseIf (curtool = T_selection Or curtool = T_magicwand) And sel.selstate = Append Then
        If Button = vbLeftButton Then
            frmGeneral.lblToA.Caption = "Distance:"
            frmGeneral.lblTo.Caption = "X= " & 0 & " - Y= " & 0
        Else
            frmGeneral.lblTo.visible = False
            frmGeneral.lblToA.visible = False
        End If
    ElseIf curtool = t_spline Then
        'frmGeneral.Label6.Visible = True
        'frmGeneral.Label6.Caption = lastsplinesizex & " " & lastsplinesizey
        If Button = vbLeftButton Then
            If Not SharedVar.splineInProgress Then
                lastsplinesizex = 1
                lastsplinesizey = 1
            Else
                lastsplinesizex = Abs(tempx - fromtilex) + 1
                lastsplinesizey = Abs(tempy - fromtiley) + 1
            End If
        End If
        frmGeneral.lblToA.Caption = "Size:"
        frmGeneral.lblFrom = "X= " & fromtilex & " - Y= " & fromtiley
        frmGeneral.lblTo.Caption = lastsplinesizex & " " & Chr(215) & " " & lastsplinesizey

    ElseIf curtool = T_tiletext Then
        If Button = vbLeftButton Then
            frmGeneral.lblFrom = "X= " & fromtilex & " - Y= " & fromtiley
        Else
            frmGeneral.lblTo.visible = False
            frmGeneral.lblToA.visible = False
        End If
    Else
        frmGeneral.lblToA.Caption = "Size:"
        frmGeneral.lblTo.Caption = "1 " & Chr(215) & " 1"
    End If

    frmGeneral.lblFrom.visible = True
    frmGeneral.lblFromA.visible = True
    If curtool <> T_bucket And curtool <> T_magnifier And curtool <> T_hand And curtool <> T_tiletext Then
        frmGeneral.lblTo.visible = True
        frmGeneral.lblToA.visible = True
    Else
        frmGeneral.lblTo.visible = False
        frmGeneral.lblToA.visible = False
    End If

End Sub

Private Sub updateCoordLabels(Button As Integer, curtilex As Integer, curtiley As Integer)
    If curtool = T_airbrush Or curtool = T_dropper Or curtool = T_Eraser Or curtool = T_pencil Or _
       curtool = T_hand Then
        If frmGeneral.lblTo.Caption <> "X= " & curtilex - fromtilex & " - Y= " & curtiley - fromtiley Then
            frmGeneral.lblTo.Caption = "X= " & curtilex - fromtilex & " - Y= " & curtiley - fromtiley
        End If
    ElseIf curtool = T_bucket Or curtool = T_magnifier Or curtool = T_tiletext Then
        'Do nothing
    ElseIf Button And (curtool = T_selection Or curtool = T_magicwand) And sel.selstate = Append Then
        If frmGeneral.lblTo.Caption <> "X= " & curtilex - fromtilex & " - Y= " & curtiley - fromtiley Then
            frmGeneral.lblTo.Caption = "X= " & curtilex - fromtilex & " - Y= " & curtiley - fromtiley
        End If
    ElseIf (Button And curtool <> t_spline) Or SharedVar.splineInProgress Then
        If usingctrl Then
            If frmGeneral.lblTo.Caption <> Abs(tempx - fromtilex) * 2 + 1 & " " & Chr(215) & " " & Abs(tempy - fromtiley) * 2 + 1 Then
                frmGeneral.lblTo.Caption = Abs(tempx - fromtilex) * 2 + 1 & " " & Chr(215) & " " & Abs(tempy - fromtiley) * 2 + 1
            End If
            If frmGeneral.lblFrom.Caption <> "X= " & fromtilex - (Abs(tempx - fromtilex)) & " - Y= " & fromtiley - (Abs(tempy - fromtiley)) Then
                frmGeneral.lblFrom.Caption = "X= " & fromtilex - (Abs(tempx - fromtilex)) & " - Y= " & fromtiley - (Abs(tempy - fromtiley))
            End If
        Else
            If frmGeneral.lblTo.Caption <> Abs(tempx - fromtilex) + 1 & " " & Chr(215) & " " & Abs(tempy - fromtiley) + 1 Then
                frmGeneral.lblTo.Caption = Abs(tempx - fromtilex) + 1 & " " & Chr(215) & " " & Abs(tempy - fromtiley) + 1
            End If
            If frmGeneral.lblFrom.Caption <> "X= " & fromtilex & " - Y= " & fromtiley Then
                frmGeneral.lblFrom.Caption = "X= " & fromtilex & " - Y= " & fromtiley
            End If
        End If
    End If

End Sub

Sub DrawWallTilesPreview()
    Call frmGeneral.cTileset.DrawWalltilesTileset
End Sub



Private Sub PlaceCursor(X As Single, Y As Single)
    'Places cursor in preview and draws tile preview if needed
    
    Dim PreviewSelection As TilesetSelection
    Dim previewOptions   As DrawOptions
    
    PreviewSelection = tileset.selection(tileset.lastButton)
    
    'last position and size of the cursor
    Static lastpos As Coordinate  'Last position of the cursor
    Static lastsize As Coordinate 'Last used cursor size (in pixels)
    
    
    'initialize to non-zero values ''' Why did I put that?... pft.
'20        If lastsize.X = 0 Then
'30            lastsize.X = 1
'40            lastsize.Y = 1
'50        End If
    
    previewOptions.drawshape = DS_Rectangle
    
    Dim cursizepixel As Coordinate 'Size in pixels of the cursor
    Dim cursize As Coordinate      'Size in tiles of the cursor
    Dim curpos As Coordinate
    
    
    Dim moved As Boolean


    
    If curtool = T_tiletext And TileText.isActive Then Exit Sub


  previewOptions.size = 1

    If curtool = T_Eraser Or curtool = T_replacebrush Then
        previewOptions.size = frmGeneral.toolSize(curtool - 1).value
        PreviewSelection.tilenr = 0
        PreviewSelection.tileSize.X = 1
        PreviewSelection.tileSize.Y = 1
        PreviewSelection.pixelSize.X = TILEW
        PreviewSelection.pixelSize.Y = TILEW
        PreviewSelection.selectionType = TS_Tiles
        
'220           previewOptions.size = PreviewSelection.tileSize.X
        
    ElseIf curtool = T_ellipse Or curtool = T_filledellipse Or _
            curtool = T_filledrectangle Or curtool = T_line Or curtool = T_pencil Or _
            curtool = T_rectangle Or curtool = t_spline Then
          
          If PreviewSelection.tileSize.X = 1 And _
              PreviewSelection.tileSize.Y = 1 Then

              previewOptions.size = frmGeneral.toolSize(curtool - 1).value
          End If
        
    ElseIf curtool = T_airbrush Then
        PreviewSelection.tileSize.X = (frmGeneral.sldAirbSize.value * 2 + 1)
        PreviewSelection.tileSize.Y = PreviewSelection.tileSize.X
        PreviewSelection.pixelSize.X = PreviewSelection.tileSize.X * TILEW
        PreviewSelection.pixelSize.Y = PreviewSelection.tileSize.Y * TILEW
    ElseIf curtool = T_lvz Then
        'With LVZ tool, highlight the lvz image under the cursor, if any
        
'        Dim lvzidx As Integer
'        Dim moIdx As Long
'
'        If MouseDown Then
'            shpcursor.visible = False
'        ElseIf lvz.getMapObjectAtPos((X + Hscr.value) / currentzoom, (Y + Vscr.value) / currentzoom, lvzidx, moIdx) <> -1 Then
'
'            Dim mapobj As LVZMapObject
'            mapobj = lvz.getMapObject(lvzidx, moIdx)
'
'            Dim offsX As Integer, offsY As Integer
'            offsX = Hscr.value / currentzoom
'            offsY = Vscr.value / currentzoom
'            shpcursor.BorderColor = IIf(mapobj.selected, vbBlue, cursorcolor)
'            shpcursor.Left = Int((mapobj.X - offsX) * currentzoom)
'            shpcursor.Top = Int((mapobj.Y - offsY) * currentzoom)
'            shpcursor.width = Int(lvz.GetMapObjectWidth(lvzidx, moIdx) * currentzoom)
'            shpcursor.height = Int(lvz.GetMapObjectHeight(lvzidx, moIdx) * currentzoom)
'
'            shpcursor.visible = True
'        Else
            shpcursor.visible = False
'        End If
        
        Exit Sub
    Else
        If PreviewSelection.isSpecialObject Then
            PreviewSelection.isSpecialObject = False
            PreviewSelection.tilenr = 0
        End If
        PreviewSelection.tileSize.X = 1
        PreviewSelection.tileSize.Y = 1
        PreviewSelection.pixelSize.X = TILEW
        PreviewSelection.pixelSize.Y = TILEH
    End If
    
    'Now that we have the size of the tiles to draw, set cursor position and size
    
  If (previewOptions.size > 1) Then
      cursize.X = previewOptions.size
      cursize.Y = cursize.X
  Else
      cursize = PreviewSelection.tileSize
  End If
  cursizepixel.X = cursize.X * currenttilew
  cursizepixel.Y = cursize.Y * currenttilew
  
    
    curpos.X = GetCornerOfTileByScreenX(CInt(X)) - ((cursize.X - 1) \ 2) * currenttilew
    curpos.Y = GetCornerOfTileByScreenY(CInt(Y)) - ((cursize.Y - 1) \ 2) * currenttilew
    
    'curpos.X = ((X \ currenttilew) - ((cursize.X - 1) \ 2)) * currenttilew
    'curpos.Y = ((Y \ currenttilew) - ((cursize.Y - 1) \ 2)) * currenttilew
    
    
'    curpos.X = (((X + currenttilew - offsetX) \ currenttilew) - ((cursize.X - 1) \ 2)) * currenttilew
'    curpos.Y = (((Y + currenttilew - offsetY) \ currenttilew) - ((cursize.Y - 1) \ 2)) * currenttilew
    


  If (magnifier.UsingPixels And cursize.X <= 2 And cursize.Y <= 2) Or TestMap.isRunning Then
      shpcursor.visible = False
      Exit Sub
  Else
      shpcursor.visible = True
  End If
 
 

    If curtool = T_filledrectangle Or curtool = T_line Or curtool = T_pencil Or _
       curtool = T_rectangle Or curtool = t_spline Then
        If frmGeneral.optToolRound(curtool - 1).value Then
            previewOptions.drawshape = DS_Circle
        Else
            previewOptions.drawshape = DS_Rectangle
        End If
    Else
        previewOptions.drawshape = DS_Rectangle
    End If
    
'The circle cursor is kind of ugly
'550       If shpcursor.Shape <> previewOptions.drawshape Then
'560           shpcursor.Shape = previewOptions.drawshape
'570       End If
        
    moved = False

    If shpcursor.Left <> curpos.X Then
        shpcursor.Left = curpos.X
        moved = True
    End If
    If shpcursor.Top <> curpos.Y Then
        shpcursor.Top = curpos.Y
        moved = True
    End If
    If shpcursor.width <> cursizepixel.X + 1 Then
        shpcursor.width = cursizepixel.X + 1
        moved = True
    End If
    If shpcursor.height <> cursizepixel.Y + 1 Then
        shpcursor.height = cursizepixel.Y + 1
        moved = True
    End If


         
    
    '''Draw the tile preview

    If showtilepreview Then
      
      If moved And Not magnifier.UsingPixels And Not SharedVar.splineInProgress Then
        
        If curtool = T_bucket Or curtool = T_customshape Or curtool = T_ellipse Or _
            curtool = T_Eraser Or curtool = T_filledellipse Or curtool = T_filledrectangle Or _
            curtool = T_line Or curtool = T_pencil Or curtool = T_rectangle Or curtool = t_spline Then
            
'        If curtool <> T_dropper And curtool <> T_hand And curtool <> T_magnifier And _
'           curtool <> T_selection And curtool <> T_magicwand And curtool <> T_replacebrush And _
'           curtool <> T_tiletext And curtool <> T_lvz And curtool <> T_Region And _
'           curtool <> T_TestMap And curtool <> T_airbrush And curtool <> T_bucket Then
    
              If SharedVar.MouseDown = 0 Then
                    Dim curtilex As Integer, curtiley As Integer
                    
                    If PreviewSelection.selectionType = TS_Walltiles Then
                        Call UpdatePreview(False, False)
                    Else

                        
                        'redraw the area where the cursor was
                        BitBlt picPreview.hDC, lastpos.X, lastpos.Y, lastsize.X, lastsize.Y, MapLayers(DL_Buffer).hDC, lastpos.X, lastpos.Y, vbSrcCopy
                        
                        'TODO: I don't think this works with different zooms
'                        If ShowLVZ Then
''                            Call UpdateLVZPreview((lastpos.X + Hscr.value) / currentzoom, (lastpos.Y + Vscr.value) / currentzoom, (lastpos.X + lastsize.X + Hscr.value) / currentzoom, (lastpos.Y + lastsize.Y + Vscr.value) / currentzoom, False)
'                        End If
                    End If
            
  
                    curtilex = ScreenToTileX(CInt(X))
                    curtiley = ScreenToTileY(CInt(Y))
                    
  '940               If PreviewSelection.tileSize.X = 1 And PreviewSelection.tileSize.Y = 1 Then
  '950                   previewOptions.size = 1
  '960               Else
  '970                   previewOptions.size = frmGeneral.toolSize(curtool - 1).Value
  '980               End If
                
                    Call tline.SetOptions(previewOptions)
                    Call tline.SetSelection(PreviewSelection)
                    
                    Call tline.DrawLine(curtilex, curtiley, curtilex, curtiley, Nothing, True, False, False)
              End If
              
            If tileset.lastButton = vbLeftButton And shpcursor.BorderColor <> vbRed Then
              shpcursor.BorderColor = vbRed
            ElseIf tileset.lastButton = vbRightButton And shpcursor.BorderColor <> vbYellow Then
              shpcursor.BorderColor = vbYellow
            End If
    

            picPreview.Refresh
        Else
          shpcursor.BorderColor = cursorcolor
        End If
      End If
      
    Else
      shpcursor.BorderColor = cursorcolor
    End If

    lastsize = cursizepixel
    lastpos = curpos
    
End Sub

Private Sub SearchWallTiles(filename As String)
'search for walltiles matching filename
    Dim walltilepath As String

    walltilepath = GetPathTo(filename) & GetFileNameWithoutExtension(filename) & ".wtl"

    AddDebug "SearchWallTiles, Searching for: " & filename & " (walltilepath: " & walltilepath & ")"

    If FileExists(walltilepath) Then
        AddDebug "SearchWallTiles, Found walltiles " & walltilepath

        If MessageBox("Walltiles definition with the same name as the file were found, do you want to load them as well?", vbYesNo + vbQuestion, "Walltiles found") = vbYes Then
            If FileExists(walltilepath) Then
                'make sure user did not move the file while the messagebox was shown
                Call walltiles.LoadWallTiles(walltilepath)
            Else
                MessageBox walltilepath & " not found.", vbExclamation + vbOKOnly, "File not found"
            End If
        End If
    End If


End Sub

Sub StoreSettings()

    gridcolor(0) = GetSetting("GridColor0", DEFAULT_GRID_COLOR0) '13107200
    gridcolor(1) = GetSetting("GridColor1", DEFAULT_GRID_COLOR1) '200
    gridcolor(2) = GetSetting("GridColor2", DEFAULT_GRID_COLOR2) '6579300
    gridcolor(3) = GetSetting("GridColor3", DEFAULT_GRID_COLOR3) '3289650
    gridcolor(4) = GetSetting("GridColor4", DEFAULT_GRID_COLOR0)
    gridcolor(5) = GetSetting("GridColor5", DEFAULT_GRID_COLOR1)
    gridcolor(6) = GetSetting("GridColor6", DEFAULT_GRID_COLOR2)
    gridcolor(7) = GetSetting("GridColor7", DEFAULT_GRID_COLOR3)

    GridOffsetX = GetSetting("GridOffsetX", 0)
    gridblocksX = GetSetting("GridBlocksX", DEFAULT_GRID_BLOCKS)
    gridsectionsX = GetSetting("GridSectionsX", DEFAULT_GRID_SECTIONS)

    GridOffsetY = GetSetting("GridOffsetY", 0)
    gridblocksY = GetSetting("GridBlocksY", DEFAULT_GRID_BLOCKS)
    gridsectionsY = GetSetting("GridSectionsY", DEFAULT_GRID_SECTIONS)

    cursorcolor = GetSetting("CursorColor", DEFAULT_CURSOR_COLOR)
    
    regnopacity1 = GetSetting("RegnOpacity1", DEFAULT_REGNOPACITY1)
    regnopacity2 = GetSetting("RegnOpacity2", DEFAULT_REGNOPACITY2)
    
    'hide tooltip text if no coords are to be shown
    showtilecoords = GetSetting("ShowCursorCoords", 1)
    If Not showtilecoords Then picPreview.tooltiptext = ""

    showtilepreview = GetSetting("ShowPreview", 1)

    If Not showtilepreview Then shpcursor.BorderColor = cursorcolor
    
    frmGeneral.LeftTilesetColor = GetSetting("LeftColor", DEFAULT_LEFTCOLOR)
    frmGeneral.RightTilesetColor = GetSetting("LeftColor", DEFAULT_LEFTCOLOR)
    frmGeneral.TilesetBackgroundColor = GetSetting("TilesetBackground", DEFAULT_TILESETBACKGROUND)
    
    Call frmGeneral.cTileset.DrawLVZTileset(True)
    Call frmGeneral.cTileset.Redraw
    
    frmGeneral.chkForceFullRadar.value = CInt(GetSetting("ForceFullRadar", "1"))
    TimerAutosave.Enabled = CBool(GetSetting("AutoSaveEnable", "1"))
End Sub


'Adds a string to the debug log ; map name is added automatically
Sub AddDebug(str As String)
    Call mdlDebug.AddDebug(Me.Caption & " @ " & str)
End Sub

'Returns a string containing the essential information on a bitmap info header
Private Function BitmapHeaderInfoString(header As BITMAPINFOHEADER) As String
    BitmapHeaderInfoString = "Bitmap info header: " & vbCrLf _
                             & " --- Color Depth: " & header.biBitCount & vbCrLf _
                             & " --- Size: " & header.biWidth & "x" & header.biHeight & vbCrLf _
                             & " --- BiSizeImage: " & header.biSizeImage & vbCrLf _
                             & " --- Compression: " & header.biCompression
End Function


'Returns a string containing the essential information on a bitmap file header
Private Function BitmapFileInfoString(header As BITMAPFILEHEADER) As String
    BitmapFileInfoString = "Bitmap info header: " & vbCrLf _
                           & " --- bfType: " & header.bfType & vbCrLf _
                           & " --- bfSize: " & header.bfSize & vbCrLf _
                           & " --- bfReserved1: " & header.bfReserved1 & " (" & LongToUnsigned(header.bfReserved1) & ")"

End Function







Sub SetFocusAt(tileX As Integer, tileY As Integer, OffsetX As Integer, OffsetY As Integer, Optional Refresh As Boolean = False)
    If tileX >= 1024 Then tileX = 1023
    If tileY >= 1024 Then tileY = 1023
    Dim valH As Long, valV As Long
    
    valH = currenttilew * tileX - OffsetX
    
    If valH < 0 Or hScr.Max <= 1 Then
        valH = 0
    ElseIf valH > hScr.Max Then
        valH = hScr.Max

    End If

    valV = currenttilew * tileY - OffsetY
    
    If valV < 0 Or vScr.Max <= 1 Then
        valV = 0
    ElseIf valV > vScr.Max Then
        valV = vScr.Max
    End If
            


    'make a correction so the value is always divisible
    'by  TileW*zoom !
'    valH = (valH \ currenttilew) * currenttilew
'    valV = (valV \ currenttilew) * currenttilew
    If valH <> hScr.value Or valV <> vScr.value Then
        Call SetScrollbarValues(valH, valV, False)
        If Refresh Then Call UpdateLevel(True)
    End If

End Sub

Function UsingRadarFullMap() As Boolean
    If CInt(GetSetting("AutoFullMapPreview", 0)) = 1 Then
        UsingRadarFullMap = (magnifier.UsingPixels Or _
                             picPreview.width * 2 \ currenttilew >= frmGeneral.picradar.width Or _
                             picPreview.height * 2 \ currenttilew >= frmGeneral.picradar.height Or _
                             frmGeneral.chkForceFullRadar.value = vbChecked)
    Else
        UsingRadarFullMap = (frmGeneral.chkForceFullRadar.value = vbChecked)
    End If
    
End Function


Function TilesetIs8bit() As Boolean
    TilesetIs8bit = (BMPInfoHeader.biBitCount = 8)
End Function

Sub ClearSelection()
    Dim undoch As New Changes
    undoredo.ResetRedo

    Call sel.DeleteSelection(undoch, True, True)

    Call undoredo.AddToUndo(undoch, UNDO_SELECTION_CLEAR)
End Sub





Sub DoAutoSave(Optional ignoreMapchanged As Boolean = False)
    'Saves a backup copy of the map, deleting the oldest backup copy if there are too many
    'Save will only be done if mapchanged = true, or ignoreMapchanged = true
    'ignoreMapchanged is needed because frmGeneral asks for a backup copy right after it opened the map
    On Error GoTo DoAutoSave_Error
    
    Dim basepath As String
    Dim path As String
    Dim currentsaves() As String
    Dim i As Integer
    Dim f As Integer
    
    If Not mapchanged And Not ignoreMapchanged Then Exit Sub
    
    While SharedVar.MouseDown <> 0
        'wait until mouse is released to avoid interferring with user's action
        DoEvents
    Wend
        
        
    
    'create autosaves folder if needed
    
    Call CreateDir(App.path & "\DCME autosaves")
    
    i = 1
    
    basepath = App.path & "\DCME autosaves\" & eLVL.GetHashCode & "_" & Format(Now, "YYYY-MM-DD-HH-MM-SS")
    path = basepath & "_" & Me.Caption & ".bak"
    
    'if that file exists, add (i) until it doesnt exist
    'shouldnt happen because autosave delay is minimum 1 minute... but you never know
    i = 1
    While FileExists(path)
        path = basepath & "_(" & i & ")" & Me.Caption & ".bak"
        i = i + 1
    Wend
    
    'check nr of current saves, if more than MAX_AUTOSAVES, clear the oldest
    currentsaves = GetAutoSaves
    
    If currentsaves(0) <> "" Then
        If UBound(currentsaves) + 1 >= CInt(GetSetting("MaxAutosaves", DEFAULT_MAX_AUTOSAVES)) Then
            Do
                Dim bakpath As String
                'Dim logpath As String
                
                bakpath = App.path & "\DCME autosaves\" & currentsaves(0)
                'logpath = GetPathTo(bakpath) & GetFileNameWithoutExtension(bakpath) & ".log"
                
                DeleteFile bakpath
'200                   If FileExists(bakpath) Then
'210                       Kill bakpath
'220                   End If
                
'                If FileExists(logpath) Then
'                    Kill logpath
'                End If
                
                currentsaves = GetAutoSaves
            
            Loop While (UBound(currentsaves) + 1 >= GetSetting("MaxAutosaves", DEFAULT_MAX_AUTOSAVES))
        End If
    End If
    
    Call SaveMap(path, (SFdefault Or SFsilent) And Not (SFsaveLVZ))

'    'save the session log too
'
'    f = FreeFile
'
'    Open GetPathTo(path) & GetFileNameWithoutExtension(path) & ".log" For Output As #f
'    'Ecrit le texte dans le fichier
'    Print #f, frmGeneral.GetDebugLog
'    'Ferme le fichier
'    Close #f
    
    Call frmGeneral.UpdateAutoSaveList
    
    On Error GoTo 0
    Exit Sub
    
DoAutoSave_Error:
    HandleError Err, "frmMain.DoAutoSave"
End Sub

'Counts the number of autosaves currently available corresponding to the current map
Function GetAutoSaves() As String()
    Dim AutoSavesCount As Long
    Dim filename As String
    Dim ret() As String
    
    AutoSavesCount = 0
    
    filename = Dir$(App.path & "\DCME autosaves\" & eLVL.GetHashCode & "*.bak")
    If filename <> "" Then
        'at least one autosave was found
        
        While filename <> ""
        
            ReDim Preserve ret(AutoSavesCount) As String
            
            
            ret(AutoSavesCount) = filename
            AutoSavesCount = AutoSavesCount + 1
            
            filename = Dir$()
        
        Wend
    Else
        ReDim ret(0)
        ret(0) = ""
    End If
    GetAutoSaves = ret
End Function


'attempts tiledata recovery, returns corrected tiledata start position
'returns a negative number if all errors cannot be avoided
'Function TileDataRecovery(startSearch As Long, f As Integer, Optional maxAttempts As Long = 24) As Long
'    Const MAX_ERRORS = 2000
'
'    Dim errorCount As Long
'    Dim nrtiles As Long
'    Dim bestOffset As Long
'    Dim bestOffsetErrorCount As Long
'    bestOffsetErrorCount = MAX_ERRORS
'    Dim bestOffsetTilesCount As Long
'    Dim offset As Long
'
'    AddDebug "TileDataRecovery, attempting tile data recovery at " & startSearch
'
'    offset = 0
'    While (offset < maxAttempts And bestOffsetErrorCount > 0)
'  Seek #f, startSearch + offset
'
'  nrtiles = 0
'  errorCount = 0
'  'keep em coming till we have reached the end of the file
'  Do Until EOF(f)
'      Dim X As Long
'      Dim Y As Long
'      Dim b(3) As Byte
'      Dim tilenr As Long
'
'      'retrieve 4 bytes
'      Get #f, , b
'
'      'extract the data
'      X = b(0) + 256 * (b(1) Mod 16)
'      Y = b(1) \ 16 + 16 * b(2)
'      tilenr = CLng(b(3))
'
'      If X < 0 Or X > 1023 Or Y < 0 Or Y > 1023 Or tilenr < 0 Or tilenr > 255 Then
'
'          errorCount = errorCount + 1
'
'      ElseIf tilenr > 0 Then
'          nrtiles = nrtiles + 1
'      End If
'
'      If errorCount > MAX_ERRORS Then Exit Do
'
'  Loop
'
'  If offset = 0 Or errorCount < bestOffsetErrorCount Then
'      bestOffset = offset
'      bestOffsetErrorCount = errorCount
'      bestOffsetTilesCount = nrtiles
'  End If
'
'  AddDebug "TileDataRecovery, with offset " & offset & ", " & nrtiles & " tiles found, " & errorCount & " errors occured."
'  offset = offset + 1
'    Wend
'
'    AddDebug "TileDataRecovery, best offset found: " & bestOffset & " with " & bestOffsetErrorCount & " errors."
'
'    If bestOffsetErrorCount > 0 Then
'        TileDataRecovery = -(startSearch + bestOffset)
'    Else
'        TileDataRecovery = startSearch + bestOffset
'    End If
'End Function

Sub Form_Resize_Force()
    Call Form_Resize
End Sub

Private Sub Vscr_Scroll()
    Call Vscr_Change
    
    'put focus back on picpreview
    On Error Resume Next
    If Not TestMap.isRunning Then picPreview.setfocus
End Sub






Sub ExportLayers()
    frmGeneral.pictemp.width = picPreview.width
    frmGeneral.pictemp.height = picPreview.height
    
    Dim i As Integer
    
    For i = DL_Regions To DL_Buffer
        frmGeneral.pictemp.Cls
        
        BitBlt frmGeneral.pictemp.hDC, 0, 0, frmGeneral.pictemp.width, frmGeneral.pictemp.height, MapLayers(i).hDC, 0, 0, vbSrcCopy
    
        frmGeneral.pictemp.Picture = frmGeneral.pictemp.Image

        Call SavePicture(frmGeneral.pictemp.Picture, App.path & "\Layer" & i & ".bmp")
    Next
End Sub

Public Property Get cursorcolor() As Long

    cursorcolor = m_lcursorcolor

End Property

Public Property Let cursorcolor(ByVal lcursorcolor As Long)

    m_lcursorcolor = lcursorcolor

End Property
