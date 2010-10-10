Attribute VB_Name = "Constants"
Option Explicit

Global Const C_TRUE = 1
Global Const C_FALSE = 0

Global Const ERR_MSINET_REQUESTTIMEDOUT = 35761

Global Const ERR_UNINITIALIZED_LVZ_DC = 20000

Global Const PI = 3.14159265

Global Const ImageTilesetW = 32 'Width of the preview of an image in the LVZ library
Global Const ImageTilesetH = 32

Global Const TILEW = 16  'Tile width in pixels
Global Const TILEH = 16  'Tile height in pixels
Global Const MAPW = 1024 'Map width in tiles
Global Const MAPH = 1024 'Map height in tiles
Global Const MAPWpx = 16384 'Map width in pixels
Global Const MAPHpx = 16384 'Map height in pixels

Global Const WT_NTILESW = 4   'Number of tiles in a walltile set (horizontally)
Global Const WT_NTILESH = 4   'Number of tiles in a walltile set (vertically)

Global Const WT_SETW = 64 'Width of a walltile set in pixels
Global Const WT_SETH = 64 'Height of a walltile set in pixels


Global Const MAX_AIRBR_SIZE = 128
Global Const MAX_AIRBR_DENSITY = 200

Global Const MAX_TE_AIRBR_RAD = 128
Global Const MAX_TE_TOOLWIDTH = 128

Global Const BUCKET_STACK = 50000
Global Const MAX_BUCKETFILL = 128000
Global Const MAX_MAGICWAND = 128000

'Tile Editor Form Size
Global Const MAX_TE_FORMWIDTH = 320
Global Const MAX_TE_FORMHEIGHT = 438

'increase for slower map dragging while drawing
Global Const DRAG_FACTOR_LEFT = 8
Global Const DRAG_FACTOR_RIGHT = 12
Global Const DRAG_FACTOR_TOP = 12
Global Const DRAG_FACTOR_BOTTOM = 8

'index of the region tool buttons
Global Const REGION_MAGICWAND = 0
Global Const REGION_SELECTION = 1

'default settings
Global Const DEFAULT_AUTOSAVE_DELAY = 10
Global Const DEFAULT_MAX_AUTOSAVES = 5


Global Const DEFAULT_GRID_BLOCKS = 16
Global Const DEFAULT_GRID_SECTIONS = 4

Global Const GREY30 = 1973790
Global Const GREY50 = 3289650
Global Const GREY80 = 5263440
Global Const GREY100 = 6579300
Global Const GREY130 = 8553090
Global Const GREY200 = 13158600

Global Const DEFAULT_GRID_COLOR0 = vbCyan
Global Const DEFAULT_GRID_COLOR1 = 8388608  'RGB(0,0,128)
Global Const DEFAULT_GRID_COLOR2 = GREY100 'RGB(100, 100, 100)
Global Const DEFAULT_GRID_COLOR3 = GREY30  'RGB(50, 50, 50)

Global Const DEFAULT_CURSOR_COLOR = GREY200

Global Const DEFAULT_REGNOPACITY1 = 90
Global Const DEFAULT_REGNOPACITY2 = 50

Global Const DEFAULT_LEFTCOLOR = vbRed
Global Const DEFAULT_RIGHTCOLOR = vbYellow
Global Const DEFAULT_TILESETBACKGROUND = vbGreen


Global Const DEFAULT_IMAGEEDITOR = "mspaint"


Global Const RADAR_OUTSIDE_COLOR = GREY80




'UNDO ACTION DESCRIPTORS
Global Const UNDO_SHIPDRAW = "Draw Tiles With Ship"
Global Const UNDO_SWITCH = "Switch Tiles"
Global Const UNDO_REPLACE = "Replace Tiles"
Global Const UNDO_SELECTION_SWITCH = "Switch Tiles In Selection"
Global Const UNDO_SELECTION_REPLACE = "Replace Tiles In Selection"
Global Const UNDO_SELECTNONE = "Apply Selection"
Global Const UNDO_SELECTNONE_INSCREEN = "Unselect Tiles In Screen"
Global Const UNDO_SELECTION_MIRROR = "Flip Selection Horizontally"
Global Const UNDO_SELECTION_FLIP = "Flip Selection Vertically"
Global Const UNDO_SELECTION_ROTATE90 = "Rotate Selection 90° Right"
Global Const UNDO_SELECTION_ROTATE180 = "Rotate Selection 180°"
Global Const UNDO_SELECTION_ROTATE270 = "Rotate Selection 90° Left"
Global Const UNDO_SELECTION_ROTATEFREE = "Rotate Selection"
Global Const UNDO_SELECTION_APPLY = "Apply Selection"
Global Const UNDO_SELECTION_MOVE = "Move Selection"
Global Const UNDO_SELECTION_REMOVEAREA = "Remove Tiles From Selection"
Global Const UNDO_SELECTION_ADDAREA = "Add Tiles To Selection"

Global Const UNDO_WAND_ADDAREA = "Add Tiles To Selection"
Global Const UNDO_WAND_REMOVEAREA = "Remove Tiles From Selection"
Global Const UNDO_WAND_APPLY = "Apply Selection"
Global Const UNDO_WAND_APPLY_AND_ADD = "New Selection"
Global Const UNDO_WAND_MOVE = "Move Selection"

Global Const UNDO_SELECTION_APPLY_AND_ADD = "New Selection"
Global Const UNDO_SELECTION_CLEAR = "Clear Selection"
Global Const UNDO_SELECTION_PASTE = "Paste" ' (SetSelectionData)
Global Const UNDO_SELECTION_RESIZE = "Resize Selection"
Global Const UNDO_SELECTION_TOWALLTILE = "Convert Selection To Wall Tiles"
Global Const UNDO_SELECTALLTILES = "Select Tiles"
Global Const UNDO_SELECTTILENR = "Select Tile #" '( & "  " & tilenr )
Global Const UNDO_UNSELECTTILENR = "Unselect Tile #"
Global Const UNDO_PENCIL = "Pencil"
Global Const UNDO_BUCKETFILL = "Bucket Fill"
Global Const UNDO_SPLINE = "Polyline"
Global Const UNDO_AIRBRUSH = "Airbrush"

Global Const UNDO_REPLACEBRUSH = "ReplaceBrush"
'Global Const UNDO_LINE = "Line" 'it uses Toolname(curtool)
Global Const UNDO_TILETEXT = "Tile Text"

Global Const UNDO_TEXTTOMAP = "Text To Map"

Global Const UNDO_PICTOMAP = "Picture To Map"
Global Const UNDO_ERASER = "Eraser"

Global Const UNDO_REGION_ADD = "Add Tiles To Region"
Global Const UNDO_REGION_REMOVE = "Remove Tiles From Region"
Global Const UNDO_REGION_NEW = "Create New Region"
Global Const UNDO_REGION_CHANGEPROPERTIES = "Change Region Properties"
    
    
'Global Const DEFAULT_UPDATE_URL = "http://www.student.kuleuven.ac.be/~s0158884/DCME/dcmeupdate.txt"
Global Const DEFAULT_UPDATE_URL = "http://www.dcme.sscentral.com/autoupdate/dcmeupdate.txt"



