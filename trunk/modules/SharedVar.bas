Attribute VB_Name = "SharedVar"
Option Explicit


'true if mouse if down on tileset/map/radar
Global MouseDown As Integer

Global bDEBUG As Boolean

Global generalLoadStarted As Boolean    'frmGeneral has started to load
Global generalLoaded As Boolean         'frmGeneral has finished loading
Global generalHwnd As Long

Global inSplash As Boolean


'Global screenUpdating As Boolean

'Global tick As Long

Global clipdata() As Integer
Global clipBoundaries As area
Global clipBitField As boolArray
Global clipHasData As Boolean

Type DrawArea
    X As Long
    Y As Long
    width As Long
    height As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type area
    Left As Integer
    Right As Integer
    Top As Integer
    Bottom As Integer
End Type




Enum TilesetTabs
    'Should be indexes of the corresponding tabs
    
    TB_Tiles = 1
    TB_Walltiles
    TB_LVZ
    'TB_Block
End Enum

Type Coordinate
    X As Long
    Y As Long
End Type

Type CoordinateLong
    X As Long
    Y As Long
End Type

Enum TilesetSelectionType
    TS_Tiles = 1
    TS_Walltiles = 2
    TS_LVZ = 3
    TS_Block = 4
End Enum

'represents a single tileset selection
Type TilesetSelection
    selectionType As TilesetSelectionType
        
    'for tiles
    tilenr As Integer
    tileSize As Coordinate
    
    isSpecialObject As Boolean
    
        'for lvz images and walltiles
    group As Integer
        'not used for tileset
        'gives walltile set for walltiles
        'gives lvz index for lvz images
            
        'for lvz images
    pixelSize As CoordinateLong     'size in pixels
End Type

Enum DrawShapes
    DS_Rectangle = 0
    DS_Circle = 3
End Enum

Public Type DrawOptions
    size As Integer
    step As Integer
    drawshape As DrawShapes
    Filled As Boolean
    Diagonal As Boolean
    IgnoreThickness As Boolean
    inScreen As Boolean
    
    TeethSize As Integer
    TeethNumber As Integer
    Density As Integer
    
End Type

Global Const TOOLCOUNT = 22
Enum toolenum
    T_magnifier = 1
    T_selection
    T_magicwand
    T_freehandselection
    T_hand
    T_pencil
    T_dropper
    T_Eraser
    T_airbrush
    T_replacebrush
    T_bucket
    T_line
    T_spline
    T_rectangle
    T_ellipse
    T_filledrectangle
    T_filledellipse
    T_customshape
    T_tiletext
    T_Region
    T_TestMap
    T_lvz
End Enum



Enum customshapeEnum
    s_cogwheel = 0
    s_star = 1
    s_regular = 2
End Enum


Enum saveFlags

    
    '''PART OF DEFAULT
    SFsaveExtraTiles = &H1&
    SFsaveLVZ = &H2&
    
    SFsaveTileset = &H4&
'    SFsaveRevert = &H8&
    
    
    SFsaveELVL = &HBF00& 'includes all elvl
    SFsaveELVLattr = &H100&
    SFsaveELVLregn = &H200&
    SFsaveELVLdcwt = &H400&
    SFsaveELVLdctt = &H800&
    SFsaveELVLdclv = &H1000&
    SFsaveELVLdcbm = &H2000&
    '
    SFsaveELVLunknown = &H8000&
    
    
    SFdefault = &HFFFF&
    SFoptimized = SFsaveExtraTiles Or SFsaveTileset Or SFsaveELVLattr Or SFsaveELVLunknown Or SFsaveELVLregn Or SFsaveLVZ
    
    '''NOT PART OF DEFAULT
    SFsilent = &H10000
End Enum

Enum enumPasteType
    p_normal = 1 'normal... duh
    p_under = 2 'under; can only drop/draw tiles in empty space
    p_trans = 3 'transparent; empty tiles ignored
End Enum

' THIS IS A MAP REGION, NOT HOW IT IS DEFINED IN THE ELVL STRUCTURE
' THIS IS USED IN TO QUICKLY ACCESS REGION DATA AND NEEDS
' TO BE CONVERTED TO REGIONS IN ELVL FORMAT

Type Unknown_chunk
    Type As String
    size As Long
    Data() As Byte
End Type

'Type MAPregionType
'    bitfield As New boolArray
'    name As String
'    isBase As Boolean
'    isNOAntiWarp As Boolean
'    isNOWeapon As Boolean
'    isNOFlagDrop As Boolean
'    isAutoWarp As Boolean
'    autowarpX As Integer
'    autowarpY As Integer
'    autowarpArena As String
'    pythonCode As String
'    color As Long
'
'    'data that is not saved
'    visible As Boolean
'
'    unknownChunk() As Unknown_chunk
'    unknownCount As Long
'End Type

'''''''''''''''''''''''
'TESTMAP-related types'
'''''''''''''''''''''''
Type ShipStats
    vx As Long
    vy As Long

    ship As ShipTypeEnum
    X As Double
    Y As Double

    turbo As Boolean

    energy As Double

    aimangle As Double

    freq As Integer
End Type

Type ShipProperties
    InitialThrust As Double
    MaximumThrust As Double
    InitialSpeed As Long
    MaximumSpeed As Long

    MaximumEnergy As Long
    Recharge As Double

    BombSpeed As Long
    Rotation As Double

    radius As Double

    Xsize As Integer
    Ysize As Integer
    
    BombThrust As Double
    BombFireDelay As Long
    
    BulletFireDelay As Long
    MultiFireDelay As Long
    
    BulletFireEnergy As Long
    MultiFireEnergy As Long
    BombFireEnergy As Long
    BombFireEnergyUpgrade As Long
    
    BombInitialLevel As Integer
    BombMaximumLevel As Integer
    BulletInitialLevel As Integer
    BulletMaximumLevel As Integer
    AfterBurnerEnergy As Double
End Type

Type MapSettings
    BounceFactor As Double

    SpawnX(3) As Integer
    SpawnY(3) As Integer
    SpawnRadius(3) As Integer
End Type

Enum ShipTypeEnum
    shpWarbird
    shpJavelin
    shpSpider
    shpLeviathan
    shpTerrier
    shpWeasel
    shpLancaster
    shpShark
End Enum

''''''''''''''''''''''
'UNDO-related types  '
''''''''''''''''''''''
Type typeUNDOACTION
    ChgType As enumCHANGETYPE
    ChgData() As Byte
End Type

Type typeUNDOTILECHANGE
    X As Integer
    Y As Integer
    tilenr As Integer
End Type

Enum enumCHANGETYPE
    'map-related
    MapTileChange = 1   'map tile changed

    SelTileChange   'tile changed within selection
    SelAdd          'map tile added to selection
    SelDrop         'tile dropped from selection to map
    SelMove         'selection moved
    SelDelete       'tile cleared from sel without being dropped
    SelNew          'new tile created in selection (pasting, for example)

    SelFlip         'selection was flipped vertically
    SelMirror       'selection was flipped horizontally
    SelRotateCW     'selection was rotated 90 degrees
    SelRotateCCW    'selection was rotated 270 degrees
    SelRotate180    'selection was rotated 180 degrees

    SelPaste        'selection was pasted

    'region-related
    RegionAdd       'tile added to region
    RegionRemove    'tile removed from region
    RegionNew       'new region created
    RegionDelete    'region deleted
    RegionRename    'region renamed
    RegionProperties    'region properties changed

    'lvz-related

    'misc
End Enum

''''''''''''''''''''''
'LVZ-related types   '
''''''''''''''''''''''
Enum LVZFileTypeEnum
    lvz_image = 1
    lvz_sound = 2
    lvz_other = 4
End Enum
    
Enum LVZLayerEnum
    lyr_BelowAll
    lyr_AfterBackground
    lyr_AfterTiles
    lyr_AfterWeapons
    lyr_AfterShips
    lyr_AfterGauges
    lyr_AfterChat
    lyr_TopMost
End Enum

Enum LVZModeEnum
    md_ShowAlways
    md_EnterZone
    md_EnterArena
    md_Kill
    md_Death
    md_ServerControlled
End Enum

Enum LVZScreenObjectTypes
    scr_Normal = 0
    scr_ScreenCenter_C
    scr_ScreenBottomRight_B
    scr_StatsBox_S
    scr_SpecialsTop_G
    scr_SpecialsBottom_F
    scr_EnergyBar_E
    scr_Chat_T
    scr_Radar_R
    scr_Clock_O
    scr_WeaponsTop_W
    scr_WeaponsBottom_V
    
    ' 1 = C - Screen center
    ' 2 = B - Bottom right corner
    ' 3 = S - Stats box, lower right corner
    ' 4 = G - Top right corner of specials
    ' 5 = F - Bottom right corner of specials
    ' 6 = E - Below energy bar & spec data
    ' 7 = T - Top left corner of chat
    ' 8 = R - Top left corner of radar
    ' 9 = O - Top left corner of radar's text (clock/location)
    '10 = W - Top left corner of weapons
    '11 = V - Bottom left corner of weapons
End Enum



Type LVZImageDefinition
    animationFramesX As Integer 'i16 - how many columns in the animation
    animationFramesY As Integer 'i16 - how many rows in the animation
    animationTime As Integer 'i16 - How long does the whole animation lasts.
                             'NOTE: This is stored in 1/100th of a second not 1/10
                             ' uh?? Not true -> 'u16  - time for a single frame, in 1/100th of a second
    imagename As String
    picboxIdx As Integer 'index of the picturebox in which is stored the image
    
    
    picWidth As Integer
    picHeight As Integer
    
    
    CurrentFrame As Long
    lastFrameChange As Long
End Type



Type LVZMapObject
    X As Integer
    Y As Integer
    
    'used to indicate drawing order
'    ZBuffer As Long
    
    imgidx As Integer
    layer As LVZLayerEnum
    mode As LVZModeEnum
    displayTime As Long 'u12 Display Time    How long will display for, in 1/10th of a second.
    objectID As Integer
    
    selected As Boolean
End Type


Type MapObjectRef
    lvzidx As Integer
    objidx As Long
End Type


Type LVZScreenObject
    X As Integer
    Y As Integer
    typeX As LVZScreenObjectTypes
    typeY As LVZScreenObjectTypes
    imgidx As Integer
    layer As LVZLayerEnum
    mode As LVZModeEnum
    displayTime As Long 'u12 Display Time    How long will display for, in 1/10th of a second.
    objectID As Integer
End Type

Enum LVZFilePurpose
    lvz_otherPurpose
    lvz_MapObject
    lvz_ScreenObject
End Enum

Type LVZFileStruct
    path As String
    Type As LVZFileTypeEnum
    purpose As LVZFilePurpose
End Type

Enum ContinuumLevelObjectTypes
    Invalid
    CLV1
    CLV2
End Enum




Type LVZheader
    filetype(3) As Byte '4-len string, should be "CONT"
    size(3) As Byte 'u32 , either the number of compressed sections,
                    '      or the decompressed data size
End Type

Type CLVheader
    filetype(3) As Byte '4-len string, should be "CLV1" or "CLV2"
    objcount(3) As Byte 'u32 , How many object (mapobject and screenobject) definitions are in this Object Section.
                    '      or the decompressed data size
    imgCount(3) As Byte 'How many image (imageobject) definitions are in this Object Section.
End Type




Type LVZstruct
    name As String
    path As String
    files() As LVZFileStruct
    mapobjects() As LVZMapObject
    
    'Indicates the index of the next mapobject to insert on each layer
    nextIndexFor(lyr_BelowAll To lyr_TopMost) As Long
    
    screenobjects() As LVZScreenObject
    imagedefinitions() As LVZImageDefinition
    filecount As Long
    mapObjectCount As Long
    ScreenobjectCount As Long
    ImageDefinitionCount As Integer
    
    totalSelected As Long 'Total number of selected objects from this lvz
End Type


''''''''''''''''''''''
'UPDATE-related types'
''''''''''''''''''''''

Type tFileToUpdate
    url As String
    localpath As String
    description As String
    version As Long
    sfx As Boolean
    DCMEexe As Boolean
    filesize As Long
End Type

Type tChange
    description As String
    version As Long
End Type


Global splineInProgress As Boolean

Global curtool As toolenum
Global curCustomShape As customshapeEnum

Global MAX_TOOL_SIZE(16) As Integer



Function ToolName(ByVal t As toolenum) As String
    Select Case t
    Case T_magnifier
        ToolName = "Magnifier"
    Case T_selection
        ToolName = "Selection"
    Case T_magicwand
        ToolName = "MagicWand"
    Case T_freehandselection
        ToolName = "FreehandSelection"
    Case T_hand
        ToolName = "Hand"
    Case T_pencil
        ToolName = "Pencil"
    Case T_dropper
        ToolName = "Dropper"
    Case T_Eraser
        ToolName = "Eraser"
    Case T_airbrush
        ToolName = "AirBrush"
    Case T_replacebrush
        ToolName = "ReplaceBrush"
    Case T_bucket
        ToolName = "Bucket"
    Case T_line
        ToolName = "Line"
    Case T_spline
        ToolName = "SpLine"
    Case T_rectangle
        ToolName = "Rectangle"
    Case T_ellipse
        ToolName = "Ellipse"
    Case T_filledrectangle
        ToolName = "FilledRectangle"
    Case T_filledellipse
        ToolName = "FilledEllipse"
    Case T_customshape
        ToolName = CustomShapeName(curCustomShape)

        ToolName = "TileText"
    Case T_Region
        ToolName = "Regions"
    Case T_TestMap
        ToolName = "TestMap"
    Case T_lvz
        ToolName = "LVZ"
    Case Else
        ToolName = ""
    End Select
End Function

'Function GetToolFromName(sname As String) As toolenum
'    Dim i As toolenum
'
'    If Len(sname) > 0 Then
'        For i = 1 To TOOLCOUNT
'            If ToolName(i) = sname Then
'                GetToolFromName = i
'                Exit Function
'            End If
'        Next
'    End If
'
'    GetToolFromName = T_invalid
'
'End Function

Function CustomShapeName(t As customshapeEnum) As String
    Select Case t
    Case s_cogwheel
        CustomShapeName = "Cogwheel"
    Case s_star
        CustomShapeName = "Star"
    Case s_regular
        CustomShapeName = "Polygon"
    Case Else
        CustomShapeName = ""
    End Select
End Function


Function ToolHasTilePreview(t As toolenum) As Boolean
    Select Case t
    
    Case T_bucket, T_customshape, T_ellipse, _
            T_Eraser, T_filledellipse, T_filledrectangle, _
             T_line, T_pencil, T_rectangle, T_spline
             
        ToolHasTilePreview = True
        
    Case Else
        ToolHasTilePreview = False
    End Select
End Function


'Function CustomShapeName(t As customshapeEnum) As String
'    Select Case t
'    Case 0
'        CustomShapeName = "Cogwheel"
'    Case 1
'        CustomShapeName = "Star"
'    Case 2
'        CustomShapeName = "Regular"
'    Case Else
'        CustomShapeName = ""
'    End Select
'
'End Function






Function GetRED(ByVal RGBColor As Long) As Long
    GetRED = RGBColor And &HFF&
End Function

'get the green value from a long
Function GetGREEN(ByVal RGBColor As Long) As Long
    GetGREEN = (RGBColor And &HFF00&) \ &H100&
End Function

'get the blue value from a long
Function GetBLUE(ByVal RGBColor As Long) As Long
    GetBLUE = (RGBColor And &HFF0000) \ &H10000
End Function

Function GetColor(ByRef parent As Form, oldcolor As Long, Optional fillpreview As Boolean = False, Optional showdefault As Boolean = True, Optional default As Long = vbBlack) As Long
'calls the color map form to select a color and returns the color
    Dim tmpcolor As Long
    tmpcolor = oldcolor
    
    Load frmColor
    
    Call frmColor.SetData(oldcolor, ByVal VarPtr(tmpcolor), default, showdefault, fillpreview)
    
    frmColor.show vbModal, parent
    
    'Return the modified color
    GetColor = tmpcolor
End Function


Function IsShift(Shift As Integer) As Boolean

    IsShift = (Shift Mod 2 = 1)

End Function


Function IsControl(Shift As Integer) As Boolean

    IsControl = (Shift And 2)

End Function



Function FlagIsPartial(ByVal flags As saveFlags, Check As saveFlags) As Boolean
    FlagIsPartial = ((flags And Check) <> 0)
End Function

Function FlagIs(ByVal flags As Long, Check As Long) As Boolean
    FlagIs = ((flags And Check) = Check)
End Function

Sub FlagAdd(ByRef flags As Long, add As Boolean, ByVal newflag As Long)
    If add Then
        flags = flags Or newflag
    Else
        flags = flags And Not newflag
    End If
End Sub

Function SetDefaultAnimationProperties(ByRef imgdef As LVZImageDefinition) As Boolean
    'if the file is a known continuum graphics file, sets default frame animation values
    'and returns true
    'else, sets to 1,1, 100 and returns false
    If Not IsImageType(GetExtension(imgdef.imagename)) Then
        SetDefaultAnimationProperties = False
        Exit Function
    End If
    
    SetDefaultAnimationProperties = True
        
    Dim filetitle As String
    filetitle = LCase(GetFileNameWithoutExtension(imgdef.imagename))
    
    If Len(filetitle) = 5 And Left$(filetitle, 4) = "ship" And _
        Asc(Right$(filetitle, 1)) >= Asc("1") And _
        Asc(Right$(filetitle, 1)) <= Asc("8") Then              'ships
        
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 4
        imgdef.animationTime = 600
        
    ElseIf filetitle = "ships" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 32
        imgdef.animationTime = 4800
        
    ElseIf filetitle = "over1" Then 'small asteroid 1
        imgdef.animationFramesX = 15
        imgdef.animationFramesY = 2
        imgdef.animationTime = 200
    ElseIf filetitle = "over2" Then 'big asteroid
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 3
        imgdef.animationTime = 200
    ElseIf filetitle = "over3" Then 'small asteroid 2
        imgdef.animationFramesX = 15
        imgdef.animationFramesY = 2
        imgdef.animationTime = 200
    ElseIf filetitle = "over4" Then 'space station
        imgdef.animationFramesX = 5
        imgdef.animationFramesY = 2
        imgdef.animationTime = 200
    ElseIf filetitle = "over5" Then
        imgdef.animationFramesX = 4
        imgdef.animationFramesY = 6
        imgdef.animationTime = 200
        
    ElseIf filetitle = "goal" Then
        imgdef.animationFramesX = 9
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    ElseIf filetitle = "flag" Or filetitle = "wall" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    ElseIf filetitle = "powerb" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 1
        imgdef.animationTime = 150
        
    ElseIf filetitle = "warp" Then
        imgdef.animationFramesX = 6
        imgdef.animationFramesY = 3
        imgdef.animationTime = 100
        
    ElseIf filetitle = "warppnt" Or filetitle = "prizes" Or filetitle = "super" Or filetitle = "dropflag" Or filetitle = "shield" Or filetitle = "ssshield" Or filetitle = "king" Or filetitle = "kingex" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    ElseIf filetitle = "explode0" Then 'bullet hit
        imgdef.animationFramesX = 7
        imgdef.animationFramesY = 1
        imgdef.animationTime = 50
        
    ElseIf filetitle = "explode1" Then 'ship explosion
        imgdef.animationFramesX = 6
        imgdef.animationFramesY = 6
        imgdef.animationTime = 200
        
    ElseIf filetitle = "explode2" Then 'bomb explosion
        imgdef.animationFramesX = 4
        imgdef.animationFramesY = 11
        imgdef.animationTime = 200
        
    ElseIf filetitle = "damage" Then
        imgdef.animationFramesX = 20
        imgdef.animationFramesY = 2
        imgdef.animationTime = 75
        
    ElseIf filetitle = "mines" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 8
        imgdef.animationTime = 1200
        
    ElseIf filetitle = "bombs" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 13
        imgdef.animationTime = 1950
        
    ElseIf filetitle = "empburst" Then
        imgdef.animationFramesX = 5
        imgdef.animationFramesY = 2
        imgdef.animationTime = 150
        
    ElseIf filetitle = "turret" Then
        imgdef.animationFramesX = 8
        imgdef.animationFramesY = 5
        imgdef.animationTime = 600
        
    ElseIf filetitle = "turret2" Then
        imgdef.animationFramesX = 20
        imgdef.animationFramesY = 2
        imgdef.animationTime = 600
        
    ElseIf filetitle = "spark" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    ElseIf filetitle = "bombflsh" Then
        imgdef.animationFramesX = 6
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    ElseIf filetitle = "bullets" Then
        imgdef.animationFramesX = 4
        imgdef.animationFramesY = 1
        imgdef.animationTime = 20
        
    ElseIf filetitle = "shrapnel" Then
        imgdef.animationFramesX = 10
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    ElseIf filetitle = "repel" Then
        imgdef.animationFramesX = 5
        imgdef.animationFramesY = 2
        imgdef.animationTime = 200
    
    ElseIf filetitle = "color" Or _
           filetitle = "chat" Or _
           filetitle = "gradient" Or _
           filetitle = "health" Or _
           filetitle = "radarv" Or _
           filetitle = "radarh" Or _
           filetitle = "chatbg" Then
        
        imgdef.animationFramesX = 1
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
        
    Else
        SetDefaultAnimationProperties = False
        
        imgdef.animationFramesX = 1
        imgdef.animationFramesY = 1
        imgdef.animationTime = 100
    End If
End Function

Function LVZLayerName(layer As LVZLayerEnum) As String
    Select Case layer
        Case LVZLayerEnum.lyr_BelowAll
            LVZLayerName = "BelowAll"
        Case LVZLayerEnum.lyr_AfterBackground
            LVZLayerName = "AfterBackground"
        Case LVZLayerEnum.lyr_AfterTiles
            LVZLayerName = "AfterTiles"
        Case LVZLayerEnum.lyr_AfterShips
            LVZLayerName = "AfterShips"
        Case LVZLayerEnum.lyr_AfterWeapons
            LVZLayerName = "AfterWeapons"
        Case LVZLayerEnum.lyr_AfterGauges
            LVZLayerName = "AfterGauges"
        Case LVZLayerEnum.lyr_AfterChat
            LVZLayerName = "AfterChat"
        Case LVZLayerEnum.lyr_TopMost
            LVZLayerName = "TopMost"
    End Select

End Function


Function LVZModeName(mode As LVZModeEnum) As String
    Select Case mode
        Case md_ShowAlways
            LVZModeName = "ShowAlways"
        Case md_EnterZone
            LVZModeName = "EnterZone"
        Case md_EnterArena
            LVZModeName = "EnterArena"
        Case md_Kill
            LVZModeName = "Kill"
        Case md_Death
            LVZModeName = "Death"
        Case md_ServerControlled
            LVZModeName = "ServerControlled"
    End Select
End Function

Function LVZScreenObjectPrefix(ref As LVZScreenObjectTypes) As String
    Select Case ref
    
        Case scr_Normal
            LVZScreenObjectPrefix = ""
        Case scr_ScreenCenter_C
            LVZScreenObjectPrefix = "C"
        Case scr_ScreenBottomRight_B
            LVZScreenObjectPrefix = "B"
        Case scr_StatsBox_S
            LVZScreenObjectPrefix = "S"
        Case scr_SpecialsTop_G
            LVZScreenObjectPrefix = "G"
        Case scr_SpecialsBottom_F
            LVZScreenObjectPrefix = "F"
        Case scr_EnergyBar_E
            LVZScreenObjectPrefix = "E"
        Case scr_Chat_T
            LVZScreenObjectPrefix = "T"
        Case scr_Radar_R
            LVZScreenObjectPrefix = "R"
        Case scr_Clock_O
            LVZScreenObjectPrefix = "O"
        Case scr_WeaponsTop_W
            LVZScreenObjectPrefix = "W"
        Case scr_WeaponsBottom_V
            LVZScreenObjectPrefix = "V"
    
    End Select
End Function





Function inDebug() As Boolean
    bDEBUG = True
    inDebug = True
End Function





