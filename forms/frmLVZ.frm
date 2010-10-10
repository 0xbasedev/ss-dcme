VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLVZ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LVZ Manager"
   ClientHeight    =   7440
   ClientLeft      =   330
   ClientTop       =   -45
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   44
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer TimerAnim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3000
      Top             =   6960
   End
   Begin VB.PictureBox picTabContents 
      BorderStyle     =   0  'None
      Height          =   6300
      Index           =   3
      Left            =   -75000
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   526
      TabIndex        =   11
      Top             =   480
      Width           =   7890
      Begin VB.Frame fPreview 
         Caption         =   "Preview"
         Height          =   3495
         Index           =   3
         Left            =   3240
         TabIndex        =   42
         Top             =   2520
         Width           =   4575
         Begin VB.PictureBox picPreviewAnim 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   2055
            Index           =   3
            Left            =   120
            ScaleHeight     =   137
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   289
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   4335
         End
      End
      Begin VB.Frame fAvailableScreenImages 
         Caption         =   "Available Images"
         Height          =   2415
         Left            =   0
         TabIndex        =   25
         Top             =   3600
         Width           =   3015
         Begin VB.ListBox lstScreenImgDefs 
            Height          =   2010
            ItemData        =   "frmLVZ.frx":0000
            Left            =   120
            List            =   "frmLVZ.frx":0010
            TabIndex        =   26
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame fScreenObjects 
         Caption         =   "Current Screen Objects"
         Height          =   3375
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   3015
         Begin VB.CommandButton cmdRemoveScreenObject 
            Caption         =   "&Remove Object"
            Height          =   375
            Left            =   1560
            TabIndex        =   23
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddScreenObject 
            Caption         =   "&New Object"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   1335
         End
         Begin MSComctlLib.TreeView tvScreenObjects 
            Height          =   2535
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   4471
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   317
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imglst"
            Appearance      =   1
         End
      End
      Begin DCME.PropertyList lstScreenObjectProperties 
         Height          =   2175
         Left            =   3240
         TabIndex        =   41
         Top             =   240
         Width           =   4575
         _extentx        =   9128
         _extenty        =   5530
      End
   End
   Begin VB.PictureBox picTabContents 
      BorderStyle     =   0  'None
      Height          =   6300
      Index           =   2
      Left            =   3120
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   526
      TabIndex        =   1
      Top             =   6120
      Width           =   7890
      Begin VB.Frame fPreview 
         Caption         =   "Preview"
         Height          =   3495
         Index           =   2
         Left            =   3240
         TabIndex        =   39
         Top             =   2520
         Width           =   4575
         Begin VB.PictureBox picPreviewAnim 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3135
            Index           =   2
            Left            =   120
            ScaleHeight     =   209
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   289
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   4335
         End
      End
      Begin VB.Frame fMapObjects 
         Caption         =   "Current Map Objects"
         Height          =   3375
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   3015
         Begin VB.CommandButton cmdAddMapObject 
            Caption         =   "&New Object"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdRemoveMapObject 
            Caption         =   "&Remove Object"
            Height          =   375
            Left            =   1560
            TabIndex        =   30
            Top             =   2880
            Width           =   1335
         End
         Begin MSComctlLib.TreeView tvMapObjects 
            Height          =   2535
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   4471
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   317
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imglst"
            Appearance      =   1
         End
      End
      Begin VB.Frame fAvailableMapImages 
         Caption         =   "Available Images"
         Height          =   2415
         Left            =   0
         TabIndex        =   27
         Top             =   3600
         Width           =   3015
         Begin VB.ListBox lstMapImgDefs 
            Height          =   2010
            ItemData        =   "frmLVZ.frx":0035
            Left            =   120
            List            =   "frmLVZ.frx":0045
            TabIndex        =   28
            Top             =   240
            Width           =   2775
         End
      End
      Begin DCME.PropertyList lstMapObjectProperties 
         Height          =   2175
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   4575
         _extentx        =   9128
         _extenty        =   5530
      End
   End
   Begin VB.PictureBox picTabContents 
      BorderStyle     =   0  'None
      Height          =   6300
      Index           =   1
      Left            =   -75000
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   526
      TabIndex        =   10
      Top             =   480
      Width           =   7890
      Begin VB.Frame fPreview 
         Caption         =   "Preview"
         Height          =   3495
         Index           =   1
         Left            =   3240
         TabIndex        =   37
         Top             =   2520
         Width           =   4575
         Begin VB.PictureBox picPreviewAnim 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   2055
            Index           =   1
            Left            =   120
            ScaleHeight     =   137
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   289
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   4335
         End
      End
      Begin VB.Frame fImageDefinitions 
         Caption         =   "Current Image Definitions"
         Height          =   3375
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   3015
         Begin VB.CommandButton cmdRemoveImgDef 
            Caption         =   "&Remove Image"
            Height          =   375
            Left            =   1560
            TabIndex        =   19
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddImgDef 
            Caption         =   "&New Image"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   2880
            Width           =   1335
         End
         Begin MSComctlLib.TreeView tvImageDefinitions 
            Height          =   2535
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   4471
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   317
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imglst"
            Appearance      =   1
         End
      End
      Begin VB.Frame fAvailableFiles 
         Caption         =   "Available Files"
         Height          =   2415
         Left            =   0
         TabIndex        =   16
         Top             =   3600
         Width           =   3015
         Begin MSComctlLib.TreeView tvAvailableFiles 
            Height          =   2055
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3625
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   317
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imglst"
            Appearance      =   1
         End
      End
      Begin DCME.PropertyList lstImageDefinitionProperties 
         Height          =   2175
         Left            =   3240
         TabIndex        =   36
         Top             =   240
         Width           =   4575
         _extentx        =   6376
         _extenty        =   5318
      End
   End
   Begin VB.PictureBox picTabContents 
      BorderStyle     =   0  'None
      Height          =   6300
      Index           =   0
      Left            =   240
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   526
      TabIndex        =   2
      Top             =   480
      Width           =   7890
      Begin VB.CommandButton cmdMakeIni 
         Caption         =   "Export INI..."
         Height          =   375
         Left            =   5520
         TabIndex        =   45
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame fPreview 
         Caption         =   "Preview"
         Height          =   3495
         Index           =   0
         Left            =   3240
         TabIndex        =   33
         Top             =   2520
         Width           =   4575
         Begin VB.PictureBox picPreviewAnim 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   2895
            Index           =   0
            Left            =   120
            ScaleHeight     =   193
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   289
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   4335
         End
      End
      Begin VB.Frame fLVZfiles 
         Caption         =   "Current LVZ Files"
         Height          =   5895
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   3015
         Begin MSComctlLib.TreeView tvLVZfiles 
            Height          =   5535
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   9763
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   317
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imglst"
            Appearance      =   1
            OLEDropMode     =   1
         End
      End
      Begin VB.CommandButton cmdAddLVZ 
         Caption         =   "New LVZ..."
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveLVZ 
         Caption         =   "Remove LVZ"
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add Item..."
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import LVZ..."
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export LVZ..."
         Height          =   375
         Left            =   5520
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   240
         Y1              =   48
         Y2              =   64
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   248
         Y1              =   64
         Y2              =   64
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   240
         Y1              =   120
         Y2              =   136
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   248
         Y1              =   136
         Y2              =   136
      End
      Begin VB.Line Line5 
         X1              =   352
         X2              =   352
         Y1              =   16
         Y2              =   144
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   6960
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   4440
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLVZ.frx":006A
            Key             =   "iconImage"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLVZ.frx":03BC
            Key             =   "iconSound"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLVZ.frx":05F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLVZ.frx":0993
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tblvz 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11880
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            Key             =   "Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Images"
            Key             =   "Images"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Map Objects"
            Key             =   "MapObjects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Screen Objects"
            Key             =   "ScreenObjects"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5160
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmLVZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain

Const LABEL_x = "X"
Const LABEL_y = "Y"
Const LABEL_image = "Image"
Const LABEL_layer = "Layer"
Const LABEL_mode = "Display Mode"
Const LABEL_displayTime = "Display Time (1/10th)"
Const LABEL_objectID = "Object ID"

Const LABEL_imagename = "Filename"
Const LABEL_animationFramesX = "Animation Frames X"
Const LABEL_animationFramesY = "Animation Frames Y"
Const LABEL_animationTime = "Animation Time (1/100th)"

Const LABEL_typeX = "X Relative to"
Const LABEL_typeY = "Y Relative to"

Dim previewFramesX As Integer
Dim previewFramesY As Integer
Dim previewFileName As String

Dim lvz As New LVZData

Sub setParent(frm As frmMain)
    Set parent = frm
    Call lvz.setParent(frm)
End Sub

Private Sub cmdAddImgDef_Click()
    Dim val As Integer
    
    Dim Node As Node
    Set Node = tvImageDefinitions.SelectedItem
    
    If Node Is Nothing Then Exit Sub
    
    'make sure we always have the lvz node (and not one of the files)
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
            
    'we selected an lvz
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    If lvzidx <> -1 Then
        Dim obj As LVZImageDefinition
        Call lvz.setImageDefinitionEmpty(obj)
        val = lvz.AddImageDefinitionToLVZ(lvzidx, obj)
        
        Dim nd As Node
        'Note: The actual ID of the Image
        Set nd = tvImageDefinitions.Nodes.add(Node.Text, tvwChild, lvz.getLVZname(lvzidx) & "_Image_" & val, "Image" & val & " - No File")
        nd.selected = True
        nd.parent.Expanded = True
        
        UpdateEnabledStatus
    End If
End Sub

Private Sub cmdAddItem_Click()
    On Error GoTo errh
    
    Dim i As Integer
    
    Dim Node As Node
    Set Node = tvLVZfiles.SelectedItem
    
    If Node Is Nothing Then Exit Sub
    
    'make sure we always have the lvz node (and not one of the files)
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
            
    'we selected an lvz
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    If lvzidx <> -1 Then
        cd.DialogTitle = "Add file"
        cd.flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer
        cd.Filter = "*.*|*.*"
        
        cd.ShowOpen
        
        Dim paths() As String
        paths = ExtractFilePaths(cd.filename)

        For i = 1 To UBound(paths)
            Call AddFile(paths(i), lvzidx)
        Next
        
    End If
    
Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    End If
End Sub

Private Function MakeKeyForMapObject(lvzidx As Integer, objidx As Long)
    MakeKeyForMapObject = lvz.getLVZname(lvzidx) & "_MapObject_" & objidx
End Function


Private Sub AddFile(path As String, lvzidx As Integer, Optional showpreview As Boolean = True)
    If FileExists(path) Then
        If lvz.SearchFile(path) <> "" Then
            MessageBox "There is already a file named " & GetFileTitle(path) & " in one of your lvz files.", vbExclamation + vbOKOnly, "Add file"
        ElseIf lvz.AddFileToLVZ(lvzidx, path) <> -1 Then
            Dim Node As Node
            Set Node = tvLVZfiles.Nodes.add(lvz.getLVZname(lvzidx), tvwChild, lvz.getLVZname(lvzidx) & "_" & GetFileTitle(path) & "_" & lvz.getFileCount(lvzidx) - 1, GetFileTitle(path), CInt(lvz.getLVZFileType(lvz.getFileData(lvzidx, lvz.getFileCount(lvzidx) - 1).path)))
            Node.selected = True
            Node.parent.Expanded = True
                        
            Call UpdateEnabledStatus
            If lvz.getLVZFileType(path) = lvz_image Then
                Dim previewimg As LVZImageDefinition
                previewimg.imagename = GetFileTitle(path)
                
                If Not SetDefaultAnimationProperties(previewimg) Then
                    'Add image definition
                    If lvz.AddImageDefinitionToLVZ(lvzidx, previewimg) = -1 Then
                        'Could not add (>256)
                        
                    End If
                End If
                
                If showpreview Then Call ShowPreviewFile(path, previewimg.animationFramesX, previewimg.animationFramesY, previewimg.animationTime)
                
                
                
            End If
        Else
            MessageBox "Error adding file " & GetFileTitle(path), vbExclamation + vbOKOnly, "Add file"
        End If
    Else
        MessageBox "File " & path & " not found", vbOKOnly + vbExclamation
    End If
End Sub
Private Sub cmdAddLVZ_Click()
    Dim str As String
    str = InputBox("New LVZ Name ?", , "Lvz" & lvz.getLVZCount)
    
    If str = "" Then
        Exit Sub
    End If
    
    If GetExtension(str) <> "lvz" Then str = str & ".lvz"
    
    If lvz.AddLVZ(str) Then
        Dim nd As Node
        Set nd = tvLVZfiles.Nodes.add(, , str, str, 3)
        nd.selected = True
        nd.Expanded = True
    Else
        MessageBox "Could not add lvz file, another file with the same name already exists.", vbOKOnly + vbExclamation
    End If
    
    Call UpdateEnabledStatus
End Sub

Private Sub cmdAddMapObject_Click()
    Dim val As Integer
    
    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    
    If Node Is Nothing Then Exit Sub
    
    'make sure we always have the lvz node (and not one of the files)
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
            
    'we selected an lvz
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    If lvzidx <> -1 Then
        Dim obj As LVZMapObject
        Call lvz.setMapObjectEmpty(obj)
        val = lvz.AddMapObjectToLVZ(lvzidx, obj)
        
        Dim nd As Node
        Set nd = tvMapObjects.Nodes.add(Node.Text, tvwChild, lvz.getLVZname(lvzidx) & "_MapObject_" & val, "MapObject" & val)
        nd.selected = True
        nd.parent.Expanded = True
        
        UpdateEnabledStatus
    End If
End Sub

Private Sub cmdAddScreenObject_Click()
    Dim val As Integer
    
    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    
    If Node Is Nothing Then Exit Sub
    
    'make sure we always have the lvz node (and not one of the files)
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
            
    'we selected an lvz
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    If lvzidx <> -1 Then
        Dim obj As LVZScreenObject
        Call lvz.setScreenObjectEmpty(obj)
        val = lvz.AddScreenObjectToLVZ(lvzidx, obj)
        
        Dim nd As Node
        Set nd = tvScreenObjects.Nodes.add(Node.Text, tvwChild, lvz.getLVZname(lvzidx) & "_ScreenObject_" & val, "ScreenObject" & val)
        nd.selected = True
        nd.parent.Expanded = True
        
        UpdateEnabledStatus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    On Error GoTo errh
    
    Dim Node As Node
    Set Node = tvLVZfiles.SelectedItem
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
    
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    If lvzidx = -1 Then
        'lvz not found
        Exit Sub
    End If
    
    
    cd.DialogTitle = "Export LVZ"
    cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    cd.Filter = "*.lvz|*.lvz"
    
    cd.ShowSave
    
    Call lvz.exportLVZ(cd.filename, lvzidx)
    
    Call UpdateEnabledStatus
    
Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    End If

End Sub

Private Sub cmdImport_Click()
    On Error GoTo errh
    
    cd.DialogTitle = "Import LVZ"
    cd.flags = cdlOFNHideReadOnly
    cd.Filter = "*.lvz|*.lvz"
    
    cd.ShowOpen
    
    Call lvz.importLVZ(cd.filename, False)
    
    Call RebuildFilesTree
    Call UpdateMapObjectListData
    Call RebuildMapObjectsTree
    Call UpdateEnabledStatus
    
Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    End If
End Sub

Private Sub cmdMakeIni_Click()
    On Error GoTo errh
    
    Dim Node As Node
    Set Node = tvLVZfiles.SelectedItem
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
    
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    If lvzidx = -1 Then
        'lvz not found
        Exit Sub
    End If
    
    
    cd.DialogTitle = "Export INI"
    cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    cd.Filter = "*.ini|*.ini"
    
    cd.ShowSave
    
    Call lvz.exportINI(cd.filename, lvzidx)
    
    Call UpdateEnabledStatus
    
Exit Sub
errh:
    If Err = cdlCancel Then
        Exit Sub
    End If
End Sub

Private Sub cmdOK_Click()
    
    'check if each object has an image assigned
    Dim lvzidx As Integer
    Dim obj As Long, Img As Integer
    
    For lvzidx = 0 To lvz.getLVZCount - 1
        For Img = 0 To lvz.getImageDefinitionCount(lvzidx) - 1
            If lvz.getImageDefinition(lvzidx, Img).imagename = "" Then
                Call MessageBox("Image" & Img & " from lvz '" & lvz.getLVZname(lvzidx) & "' has no file assigned!", vbExclamation)
                tblvz.Tabs("Images").selected = True
                tblvz_Click
                tvImageDefinitions.Nodes(lvz.getLVZname(lvzidx) & "_Image_" & Img).selected = True
                UpdateEnabledStatus
                UpdateImageDefinitionListData
                Exit Sub
            End If
        Next
        
        For obj = 0 To lvz.getMapobjectCount(lvzidx) - 1
            If lvz.getMapObjectImageID(lvzidx, obj) = -1 Or lvz.getMapObjectImageID(lvzidx, obj) >= lvz.getImageDefinitionCount(lvzidx) Then
                Call MessageBox("MapObject" & obj & " from lvz '" & lvz.getLVZname(lvzidx) & "' has no image assigned!", vbExclamation)
                tblvz.Tabs("MapObjects").selected = True
                tblvz_Click
                tvMapObjects.Nodes(lvz.getLVZname(lvzidx) & "_MapObject_" & obj).selected = True
                UpdateEnabledStatus
                UpdateMapObjectListData
                RefreshAvailableMapImages
                Exit Sub
            End If
        Next
        
        For obj = 0 To lvz.getScreenobjectCount(lvzidx) - 1
            If lvz.getScreenObjectImageID(lvzidx, obj) = -1 Or lvz.getScreenObjectImageID(lvzidx, obj) >= lvz.getImageDefinitionCount(lvzidx) Then
                Call MessageBox("ScreenObject" & obj & " from lvz '" & lvz.getLVZname(lvzidx) & "' has no image assigned!", vbExclamation)
                tblvz.Tabs("ScreenObjects").selected = True
                tblvz_Click
                tvScreenObjects.Nodes(lvz.getLVZname(lvzidx) & "_ScreenObject_" & obj).selected = True
                UpdateEnabledStatus
                UpdateScreenObjectListData
                RefreshAvailableScreenImages
                Exit Sub
            End If
        Next
    Next
                
    Call parent.lvz.setLVZData(lvz.getLVZData, lvz.getLVZCount)
    
    Call parent.lvz.buildAllLVZImages
    
    'Call parent.lvz.DrawLVZImageInterface(parent.lvz.curImageFrame)
'400       Call parent.tileset.DrawLVZTileset(True) <-- buildAllLVZImages already redraws the lvz tileset
    
    
    Unload Me
End Sub

Private Sub cmdRemoveImgDef_Click()
    Dim Node As Node
    Set Node = tvImageDefinitions.SelectedItem
    If Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    Dim fileIdx As Integer
    'works for filetitle's too!
    fileIdx = ParseID(Node.Key)

    If fileIdx = -1 Then
        'file not found in lvz
    Else
        Dim MapObjCount As Long, ScrObjCount As Long
        MapObjCount = lvz.CountMapObjectsUsingImage(idx, fileIdx)
        ScrObjCount = lvz.CountScreenObjectsUsingImage(idx, fileIdx)
        
        If MapObjCount > 0 Or ScrObjCount > 0 Then
            Dim ret As VbMsgBoxResult
            ret = MessageBox("There are " & MapObjCount & " map objects and " & ScrObjCount & " screen objects associated to this image. Deleting it will also delete all these objects. Do you wish to delete it?" & fileIdx & "?", vbExclamation + vbYesNo, "Confirm Delete Image Definition")
            
            If ret = vbNo Then
                Exit Sub
            End If
        End If
        

        Call lvz.RemoveLinksToImage(idx, fileIdx)
        
        Call lvz.removeImageDefinitionFromLVZ(idx, fileIdx)
        'if this doesn't work, use rebuildtree
        'Call tvLVZfiles.Nodes.Remove(Node.Index)
        Call RebuildFilesTree
        Call RebuildImageDefinitionsTree
        
        Call UpdateEnabledStatus

    End If
End Sub

Private Function ParseID(Text As String, Optional Delimiter As String = "_", Optional Index As Integer = -1) As Integer
    If Text = "" Then
        ParseID = -1
        Exit Function
    End If
    
    Dim tmp() As String
    tmp = Split(Text, Delimiter)
    
    If Index = -1 Then Index = UBound(tmp)
    
    If UBound(tmp) >= Index Then
        If IsNumeric(tmp(Index)) Then
            ParseID = CInt(tmp(Index))
        Else
            ParseID = -1
        End If
    Else
        ParseID = -1
    End If
End Function

Private Sub cmdRemoveItem_Click()
    Dim Node As Node
    Set Node = tvLVZfiles.SelectedItem

    If Node.parent Is Nothing Then
        Exit Sub
    End If
        
    
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    Dim fileIdx As Long
    'works for filetitle's too!
    fileIdx = lvz.getIndexOfFile(idx, Node.Text)
    

    
    If fileIdx = -1 Then
        'file not found in lvz
    Else
        Dim imgCount As Long
        imgCount = lvz.CountImagesUsingFile(Node.Text)
        
        If imgCount > 0 Then
            Dim ret As VbMsgBoxResult
            ret = MessageBox("There are " & imgCount & " image definitions using '" & Node.Text & "'. Deleting it will also remove all these image definitions, and all objects using them. Do you wish to delete it?", vbExclamation + vbYesNo, "Confirm Remove File")
            
            If ret = vbNo Then
                Exit Sub
            End If
        End If
        

        Call lvz.RemoveLinksToFile(Node.Text)
        
        
        Call lvz.removeFileFromLVZ(idx, fileIdx)
        'if this doesn't work, use rebuildtree
        
        Dim nodeIdx As Long
        nodeIdx = Node.Index
        
        Call RebuildFilesTree
        
        Call RestoreNodeSelection(tvLVZfiles, nodeIdx)
        
        Call tvLVZfiles_Click
        
        Call UpdateEnabledStatus

        

    End If

End Sub

Private Sub cmdRemoveLVZ_Click()
    Dim Node As Node
    Set Node = tvLVZfiles.SelectedItem
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.Text)
    
    If idx = -1 Then
        'lvz not found
    Else
        If lvz.getFileCount(idx) > 0 Then
            If MessageBox("Delete " & lvz.getLVZname(idx) & " and all its files, image definitions, map objects and screen objects?", vbYesNo + vbQuestion, "Confirm Delete LVZ") = vbNo Then
                Exit Sub
            End If
        End If
        
        Call lvz.removeLVZ(idx)
        
        Dim nodeIdx As Integer
        nodeIdx = Node.Index
        
        Call RebuildFilesTree
        
        'Call RestoreNodeSelection(tvLVZfiles, nodeIdx)
        
        Call tvLVZfiles_Click
        
        Call UpdateEnabledStatus
    End If
    
    
End Sub


Private Sub RestoreNodeSelection(ByRef tree As TreeView, nodeIdx As Long)
    Dim nodecount As Long
    nodecount = tree.Nodes.count
    
    If nodeIdx > nodecount Then
        If nodecount > 0 Then
            tree.Nodes(nodecount).selected = True
            tree.SelectedItem.EnsureVisible
        End If
    ElseIf nodeIdx > 0 Then
        tree.Nodes(nodeIdx).selected = True
        tree.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub cmdRemoveMapObject_Click()
    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    If Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    'works for filetitle's too!
    objidx = ParseID(Node.Key)
    
    
    Dim nodeIdx As Long
    'Node.Key
    nodeIdx = Node.Index
    
    Call lvz.removeMapObjectFromLVZ(idx, objidx)
    'if this doesn't work, use rebuildtree
    'Call tvMapObjects.Nodes.Remove(Node.Index)
    Call RebuildMapObjectsTree
    
    Call RestoreNodeSelection(tvMapObjects, nodeIdx)
    
    Call tvMapObjects_Click
    
    'tvMapObjects.In
End Sub

Private Sub cmdRemoveScreenObject_Click()
    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    If Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    'works for filetitle's too!
    objidx = ParseID(Node.Key)
    
    Dim nodeIdx As Long
    nodeIdx = Node.Index
    
    Call lvz.removeScreenObjectFromLVZ(idx, objidx)
    
    Call RebuildScreenObjectsTree
    
    Call RestoreNodeSelection(tvScreenObjects, nodeIdx)
    
    Call tvScreenObjects_Click
    
End Sub

Private Sub Form_Load()
   
    Set Me.Icon = frmGeneral.Icon
    
    Call lvz.setLVZData(parent.lvz.getLVZData, parent.lvz.getLVZCount)
    
    Call RebuildFilesTree
    Call RebuildMapObjectsTree
    Call RebuildScreenObjectsTree
    Call RebuildImageDefinitionsTree

    CreateImageDefinitionsProperties
    CreateScreenObjectProperties
    CreateMapObjectProperties
    
    'make sure controls are enabled correctly
    Call UpdateEnabledStatus
    
    tblvz_Click
    
End Sub

Private Sub CreateImageDefinitionsProperties()
    Call lstImageDefinitionProperties.AddProperty(LABEL_imagename, p_text)
        Call lstImageDefinitionProperties.setPropertyLocked(LABEL_imagename, True)
    Call lstImageDefinitionProperties.AddProperty(LABEL_animationFramesX, p_number, 1, 32767)
    Call lstImageDefinitionProperties.AddProperty(LABEL_animationFramesY, p_number, 1, 32767)
    Call lstImageDefinitionProperties.AddProperty(LABEL_animationTime, p_number, 1, 65535)
End Sub

Private Sub CreateScreenObjectProperties()
    Dim chlst() As String
    
    Call lstScreenObjectProperties.AddProperty(LABEL_x, p_number, -2048, 2047)
    
    Call lstScreenObjectProperties.AddProperty(LABEL_typeX, p_list)
        ReDim chlst(11)
        chlst(0) = "Screen Left"
        chlst(1) = "Screen Center"
        chlst(2) = "Screen Right"
        chlst(3) = "Stats Box Right Edge"
        chlst(4) = "Specials Right"
        chlst(5) = "Specials Right"
        chlst(6) = "Energy Bar Center"
        chlst(7) = "Chat Text Right Edge"
        chlst(8) = "Radar Left Edge"
        chlst(9) = "Clock Left Edge"
        chlst(10) = "Weapons Left"
        chlst(11) = ""
        Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_typeX, chlst)
        
    Call lstScreenObjectProperties.AddProperty(LABEL_y, p_number, -2048, 2047)
        Call lstScreenObjectProperties.AddProperty(LABEL_typeY, p_list)
        ReDim chlst(11)
        chlst(0) = "Screen Top"
        chlst(1) = "Screen Center"
        chlst(2) = "Screen Bottom"
        chlst(3) = "Stats Box Bottom Edge"
        chlst(4) = "Specials Top"
        chlst(5) = "Specials Bottom"
        chlst(6) = "Under Energy Bar"
        chlst(7) = "Chat Text Top"
        chlst(8) = "Radar Top"
        chlst(9) = "Clock Top"
        chlst(10) = "Weapons Top"
        chlst(11) = "Weapons Bottom"
        Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_typeY, chlst)
        
    Call lstScreenObjectProperties.AddProperty(LABEL_image, p_text)
        Call lstScreenObjectProperties.setPropertyLocked(LABEL_image, True)
    Call lstScreenObjectProperties.AddProperty(LABEL_layer, p_list)

        ReDim chlst(7)
        chlst(0) = "0 - Below All"
        chlst(1) = "1 - After Background"
        chlst(2) = "2 - After Tiles"
        chlst(3) = "3 - After Weapons"
        chlst(4) = "4 - After Ships"
        chlst(5) = "5 - After Gauges"
        chlst(6) = "6 - After Chat"
        chlst(7) = "7 - Top Most"
        Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_layer, chlst)
    Call lstScreenObjectProperties.AddProperty(LABEL_mode, p_list)
        ReDim chlst(5)
        chlst(0) = "0 - Show Always"
        chlst(1) = "1 - Enter Zone"
        chlst(2) = "2 - Enter Arena"
        chlst(3) = "3 - Kill"
        chlst(4) = "4 - Death"
        chlst(5) = "5 - Server Controlled"
        Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_mode, chlst)
    Call lstScreenObjectProperties.AddProperty(LABEL_displayTime, p_number)
        Call lstScreenObjectProperties.setPropertyNumberBoundaries(LABEL_displayTime, 0, 100000000)
    
    Call lstScreenObjectProperties.AddProperty(LABEL_objectID, p_number)
        Call lstScreenObjectProperties.setPropertyNumberBoundaries(LABEL_objectID, 0, 32767)
End Sub

Private Sub CreateMapObjectProperties()

    Call lstMapObjectProperties.AddProperty(LABEL_x, p_number, -32767, 32767)
    Call lstMapObjectProperties.AddProperty(LABEL_y, p_number, -32767, 32767)
    Call lstMapObjectProperties.AddProperty(LABEL_image, p_text)
        Call lstMapObjectProperties.setPropertyLocked(LABEL_image, True)
    Call lstMapObjectProperties.AddProperty(LABEL_layer, p_list)
        Dim chlst() As String
        ReDim chlst(7)
        chlst(0) = "0 - Below All"
        chlst(1) = "1 - After Background"
        chlst(2) = "2 - After Tiles"
        chlst(3) = "3 - After Weapons"
        chlst(4) = "4 - After Ships"
        chlst(5) = "5 - After Gauges"
        chlst(6) = "6 - After Chat"
        chlst(7) = "7 - Top Most"
        Call lstMapObjectProperties.setPropertyChoiceList(LABEL_layer, chlst)
    Call lstMapObjectProperties.AddProperty(LABEL_mode, p_list)
        ReDim chlst(5)
        chlst(0) = "0 - Show Always"
        chlst(1) = "1 - Enter Zone"
        chlst(2) = "2 - Enter Arena"
        chlst(3) = "3 - Kill"
        chlst(4) = "4 - Death"
        chlst(5) = "5 - Server Controlled"
        Call lstMapObjectProperties.setPropertyChoiceList(LABEL_mode, chlst)
    Call lstMapObjectProperties.AddProperty(LABEL_displayTime, p_number)
        Call lstMapObjectProperties.setPropertyNumberBoundaries(LABEL_displayTime, 0, 100000000)
    
    Call lstMapObjectProperties.AddProperty(LABEL_objectID, p_number)
        Call lstMapObjectProperties.setPropertyNumberBoundaries(LABEL_objectID, 0, 32767)
End Sub

Private Sub RebuildFilesTree()
    tvLVZfiles.Nodes.Clear
    Dim i As Integer
    Dim j As Long
    
    For i = 0 To lvz.getLVZCount - 1
        Call tvLVZfiles.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
        For j = 0 To lvz.getFileCount(i) - 1
            Call tvLVZfiles.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_" & GetFileTitle(lvz.getFileData(i, j).path) & "_" & j, GetFileTitle(lvz.getFileData(i, j).path), CInt(lvz.getLVZFileType(lvz.getFileData(i, j).path)))
        Next
    Next
End Sub

Private Sub RebuildMapObjectsTree()
    tvMapObjects.Nodes.Clear
    Dim i As Integer
    Dim j As Long
    Dim lvzname As String
    
    tvMapObjects.visible = False
    For i = 0 To lvz.getLVZCount - 1
        lvzname = lvz.getLVZname(i)
        
        Call tvMapObjects.Nodes.add(, , lvzname, lvzname, 3)
        
        If lvz.getMapobjectCount(i) > 3000 Then
            Call tvMapObjects.Nodes.add(lvzname, tvwChild, lvzname & "_MapObject_0", "<Too many objects to display (" & lvz.getMapobjectCount(i) & ")>")
        Else
            For j = 0 To lvz.getMapobjectCount(i) - 1
    '            MakeKeyForMapObject(i, j)
    '            tvMapObjects.Sorted
                Call tvMapObjects.Nodes.add(lvzname, tvwChild, lvzname & "_MapObject_" & j, "MapObject" & j)
    '            Call tvMapObjects.Nodes.add(lvzname, tvwChild, , "MapObject" & j)
            Next
        End If
    Next
    
    tvMapObjects.visible = True
End Sub

Private Sub RebuildScreenObjectsTree()
    tvScreenObjects.Nodes.Clear
    Dim i As Integer
    Dim j As Long
    
    For i = 0 To lvz.getLVZCount - 1
        Call tvScreenObjects.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
        For j = 0 To lvz.getScreenobjectCount(i) - 1
            Call tvScreenObjects.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_ScreenObject_" & j, "ScreenObject" & j)
        Next
    Next
End Sub

Private Sub RebuildImageDefinitionsTree()
    tvImageDefinitions.Nodes.Clear
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To lvz.getLVZCount - 1
        Call tvImageDefinitions.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
        For j = 0 To lvz.getImageDefinitionCount(i) - 1
            Call tvImageDefinitions.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_Image_" & j, "Image" & j & IIf(lvz.getImageDefinition(i, j).imagename <> "", " (" & lvz.getImageDefinition(i, j).imagename & ")", " - No File"))
        Next
    Next
End Sub

Private Sub RebuildAvailableFilesTree()
    tvAvailableFiles.Nodes.Clear
    Dim i As Integer
    Dim j As Long
    
    For i = 0 To lvz.getLVZCount - 1
        Dim nodeCreated As Boolean
        nodeCreated = False
        For j = 0 To lvz.getFileCount(i) - 1
            Dim tmpFiledata As LVZFileStruct
            tmpFiledata = lvz.getFileData(i, j)
            If tmpFiledata.Type = lvz_image Then
                If Not nodeCreated Then
                    Call tvAvailableFiles.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
                    nodeCreated = True
                End If
                Call tvAvailableFiles.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_Image_" & j, GetFileTitle(tmpFiledata.path))
            End If
        Next
        
    Next
End Sub

Private Sub RefreshAvailableMapImages()
    lstMapImgDefs.Clear
    
    If tvMapObjects.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    'make sure we always have the lvz node (and not one of the files)
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
    
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    Dim i As Integer
    
    For i = 0 To lvz.getImageDefinitionCount(lvzidx) - 1
        Call lstMapImgDefs.addItem("Image" & i & IIf(lvz.getImageDefinition(lvzidx, i).imagename <> "", " (" & lvz.getImageDefinition(lvzidx, i).imagename & ")", " - No File"))
    Next
End Sub

Private Sub RefreshAvailableScreenImages()
    lstScreenImgDefs.Clear

    If tvScreenObjects.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    'make sure we always have the lvz node (and not one of the files)
    If Not Node.parent Is Nothing Then
        Set Node = Node.parent
    End If
    
    Dim lvzidx As Integer
    lvzidx = lvz.getIndexOfLVZ(Node.Text)
    
    Dim i As Integer
    
    For i = 0 To lvz.getImageDefinitionCount(lvzidx) - 1
        Call lstScreenImgDefs.addItem("Image" & i & IIf(lvz.getImageDefinition(lvzidx, i).imagename <> "", " (" & lvz.getImageDefinition(lvzidx, i).imagename & ")", " - No File"))
    Next
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call lvz.setParent(Nothing)
    Set parent = Nothing
    
End Sub

Private Sub lstImageDefinitionProperties_LostFocus()
    Call lstImageDefinitionProperties.EditOff(True)
End Sub

Private Sub lstImageDefinitionProperties_PropertyChanged(propName As String)
    Dim Node As Node
    Set Node = tvImageDefinitions.SelectedItem
    If Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim imgidx As Integer
    'works for filetitle's too!
    imgidx = ParseID(Node.Key)
    
    Dim tmpImgDef As LVZImageDefinition
    tmpImgDef = lvz.getImageDefinition(idx, imgidx)
    
    Select Case propName
        Case LABEL_imagename
            tmpImgDef.imagename = val(lstImageDefinitionProperties.getPropertyValue(propName))
        Case LABEL_animationFramesX
            tmpImgDef.animationFramesX = val(lstImageDefinitionProperties.getPropertyValue(propName))
        Case LABEL_animationFramesY
            tmpImgDef.animationFramesY = val(lstImageDefinitionProperties.getPropertyValue(propName))
        Case LABEL_animationTime
            tmpImgDef.animationTime = val(lstImageDefinitionProperties.getPropertyValue(propName))
    End Select
    
    Call lvz.setImageDefinition(idx, imgidx, tmpImgDef)
    
    Dim path As String
    path = lvz.SearchFile(tmpImgDef.imagename)
    
    If FileExists(path) Then
        Call showpreview(tmpImgDef)
    End If
End Sub

Private Sub lstMapImgDefs_Click()
    Dim imgidx As Integer
    Dim i As Integer
    
    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    If Node Is Nothing Then Exit Sub

    Dim idx As Integer
    If Node.parent Is Nothing Then
        idx = lvz.getIndexOfLVZ(Node.Text)
    Else
        idx = lvz.getIndexOfLVZ(Node.parent.Text)
    End If
    
    For i = 0 To lstMapImgDefs.ListCount - 1
        If lstMapImgDefs.selected(i) Then
            imgidx = i
            Exit For
        End If
    Next
    
    Dim tmpImg As LVZImageDefinition
    tmpImg = lvz.getImageDefinition(idx, imgidx)
    
    Call showpreview(lvz.getImageDefinition(idx, imgidx))
End Sub

Private Sub lstMapImgDefs_DblClick()
    Dim newImgIdx As Integer
    Dim i As Integer

    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    If Node Is Nothing Then Exit Sub
    If Node.parent Is Nothing Then Exit Sub
    
    For i = 0 To lstMapImgDefs.ListCount - 1
        If lstMapImgDefs.selected(i) Then
            newImgIdx = i
            Exit For
        End If
    Next
    

    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
        
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If

    Dim objidx As Long
    'works for filetitle's too!
    objidx = ParseID(Node.Key)
    
    Call lvz.setMapObjectImageID(idx, objidx, newImgIdx)
    
    Call UpdateMapObjectListData
End Sub

Private Sub lstMapObjectProperties_LostFocus()
    Call lstMapObjectProperties.EditOff(True)
End Sub

Private Sub lstMapObjectProperties_PropertyChanged(propName As String)
    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    If Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    'works for filetitle's too!
    objidx = ParseID(Node.Key)
    
    Dim tmpMapObj As LVZMapObject
    tmpMapObj = lvz.getMapObject(idx, objidx)
    
    Select Case propName
        Case LABEL_x
            tmpMapObj.X = val(lstMapObjectProperties.getPropertyValue(propName))
        Case LABEL_y
            tmpMapObj.Y = val(lstMapObjectProperties.getPropertyValue(propName))
        Case LABEL_layer
            Dim newlayer As LVZLayerEnum
            
            newlayer = val(lstMapObjectProperties.getPropertyValue(propName))
            tmpMapObj.layer = newlayer
            
            
            'We have to keep it sorted
            objidx = lvz.ChangeMapObjectLayer(idx, objidx, newlayer)
            
            Call RebuildMapObjectsTree
            
'            Dim nodeIdx As Integer
            
            '                                       indexes start at 1, and we don't want the parent 'MapObjects' node
'            Call RestoreNodeSelection(tvMapObjects, objidx + 2)
            tvMapObjects.Nodes(MakeKeyForMapObject(idx, objidx)).selected = True
            tvMapObjects.SelectedItem.EnsureVisible
            
        Case LABEL_mode
            tmpMapObj.mode = val(lstMapObjectProperties.getPropertyValue(propName))
        Case LABEL_displayTime
            tmpMapObj.displayTime = val(lstMapObjectProperties.getPropertyValue(propName))
        Case LABEL_objectID
            tmpMapObj.objectID = val(lstMapObjectProperties.getPropertyValue(propName))
    End Select
    
    Call lvz.setMapObject(idx, objidx, tmpMapObj)
End Sub



Private Sub lstScreenImgDefs_Click()
    Dim imgidx As Integer
    Dim i As Integer
    
    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    If Node Is Nothing Then Exit Sub
    
    Dim idx As Integer
    If Node.parent Is Nothing Then
        idx = lvz.getIndexOfLVZ(Node.Text)
    Else
        idx = lvz.getIndexOfLVZ(Node.parent.Text)
    End If
    
    For i = 0 To lstScreenImgDefs.ListCount - 1
        If lstScreenImgDefs.selected(i) Then
            imgidx = i
            Exit For
        End If
    Next
    

    Call showpreview(lvz.getImageDefinition(idx, imgidx))
End Sub

Private Sub lstScreenImgDefs_DblClick()
    Dim newImgIdx As Integer
    Dim i As Integer

    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    If Node Is Nothing Then Exit Sub
    If Node.parent Is Nothing Then Exit Sub
    
    For i = 0 To lstScreenImgDefs.ListCount - 1
        If lstScreenImgDefs.selected(i) Then
            newImgIdx = i
            Exit For
        End If
    Next
    

    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
        
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If

    Dim objidx As Long
    'works for filetitle's too!
    objidx = ParseID(Node.Key)
    
    Call lvz.setScreenObjectImageID(idx, objidx, newImgIdx)
    
    Call UpdateScreenObjectListData
End Sub

Private Sub lstScreenImgDefs_Scroll()
    Call lstScreenImgDefs_Click
End Sub

Private Sub lstScreenObjectProperties_LostFocus()
    Call lstScreenObjectProperties.EditOff(True)
End Sub

Private Sub lstScreenObjectProperties_PropertyChanged(propName As String)
    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    If Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    'works for filetitle's too!
    objidx = ParseID(Node.Key)
    
    Dim tmpScrObj As LVZScreenObject
    tmpScrObj = lvz.getScreenObject(idx, objidx)
    
    Select Case propName
        Case LABEL_x
            tmpScrObj.X = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_y
            tmpScrObj.Y = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_layer
            tmpScrObj.layer = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_mode
            tmpScrObj.mode = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_displayTime
            tmpScrObj.displayTime = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_objectID
            tmpScrObj.objectID = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_typeX
            tmpScrObj.typeX = val(lstScreenObjectProperties.getPropertyValue(propName))
        Case LABEL_typeY
            tmpScrObj.typeY = val(lstScreenObjectProperties.getPropertyValue(propName))
            
    End Select
    
    Call lvz.setScreenObject(idx, objidx, tmpScrObj)
End Sub

Private Sub tblvz_BeforeClick(Cancel As Integer)
    TimerAnim.Enabled = False
    Call ClearPreview
    
    If tblvz.SelectedItem.Key = "MapObjects" Then
        Call lstMapObjectProperties.EditOff(True)
    ElseIf tblvz.SelectedItem.Key = "ScreenObjects" Then
        Call lstScreenObjectProperties.EditOff(True)
    ElseIf tblvz.SelectedItem.Key = "Images" Then
        Call lstImageDefinitionProperties.EditOff(True)
    End If
End Sub

Private Sub tblvz_Click()
    Dim i As Integer
    'show and enable the selected tblvz's controls
    'and hide and disable all others

    
    For i = 0 To tblvz.Tabs.count - 1
        If i = tblvz.SelectedItem.Index - 1 Then
            picTabContents(i).Left = tblvz.Left + 8
            picTabContents(i).Enabled = True
            picTabContents(i).visible = True
            picTabContents(i).Top = tblvz.Top + 24
            picTabContents(i).width = tblvz.width - 16
            picTabContents(i).height = tblvz.height - 32
        Else
            picTabContents(i).Left = -5000
            picTabContents(i).Enabled = False
            picTabContents(i).visible = False
        End If
    Next
    
    If tblvz.SelectedItem.Index = 1 Then
        Call RebuildFilesTree
        
    ElseIf tblvz.SelectedItem.Index = 2 Then
        Call RebuildImageDefinitionsTree
        Call RebuildAvailableFilesTree
        
    ElseIf tblvz.SelectedItem.Index = 3 Then
        Call RebuildMapObjectsTree
        Call RefreshAvailableMapImages
        
    ElseIf tblvz.SelectedItem.Index = 4 Then
        Call RebuildScreenObjectsTree
        Call RefreshAvailableScreenImages
        
    End If
    Call UpdateEnabledStatus
End Sub



Private Sub TimerAnim_Timer()
    Call RefreshPreview(False, True)
End Sub

Private Sub tvAvailableFiles_Click()
    Dim tmpfile As LVZFileStruct
    
    tmpfile = getFileFromNode(tvAvailableFiles.SelectedItem)
    
    If FileExists(tmpfile.path) Then
        If tmpfile.Type = lvz_image Then
            Call ShowPreviewFile(tmpfile.path)
        End If
    End If
End Sub

Private Sub tvAvailableFiles_NodeClick(ByVal Node As MSComctlLib.Node)
    Call tvAvailableFiles_Click
End Sub

Private Sub tvImageDefinitions_Click()
    Call UpdateImageDefinitionListData
    Call UpdateEnabledStatus
    
    Call showpreview(getImageDefinitionFromNode(tvImageDefinitions.SelectedItem))

End Sub

Private Sub tvImageDefinitions_NodeClick(ByVal Node As MSComctlLib.Node)
    Call tvImageDefinitions_Click
End Sub

Private Sub tvLVZfiles_Click()
    Call UpdateEnabledStatus
    
    Dim tmpfile As LVZFileStruct
    Dim FramesX As Integer
    Dim FramesY As Integer
    Dim animTime As Long
    Dim filetitle As String
    
    tmpfile = getFileFromNode(tvLVZfiles.SelectedItem)
    
    If FileExists(tmpfile.path) Then
        If tmpfile.Type = lvz_image Then
            'special cases for animations
            Dim previewimg As LVZImageDefinition
            
            previewimg.imagename = GetFileTitle(tmpfile.path)
            
            Call SetDefaultAnimationProperties(previewimg)
                
            Call ShowPreviewFile(tmpfile.path, previewimg.animationFramesX, previewimg.animationFramesY, previewimg.animationTime)
        End If
    End If
End Sub

Private Function getFileFromNode(Node As Node) As LVZFileStruct
    Dim ret As LVZFileStruct
    Call lvz.setFileEmpty(ret)
    
    If Node Is Nothing Then
        '...
    ElseIf Node.parent Is Nothing Then
        '...
    Else
        ret = lvz.getFileData(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
    End If
    getFileFromNode = ret
End Function

Private Function getMapObjectFromNode(Node As Node) As LVZMapObject
    Dim ret As LVZMapObject
    Call lvz.setMapObjectEmpty(ret)
    If Node Is Nothing Then
        '...
    ElseIf Node.parent Is Nothing Then
        '...
    Else
        ret = lvz.getMapObject(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
    End If
    getMapObjectFromNode = ret
End Function

Private Function getScreenObjectFromNode(Node As Node) As LVZScreenObject
    Dim ret As LVZScreenObject
    Call lvz.setScreenObjectEmpty(ret)
    If Node Is Nothing Then
        '...
    ElseIf Node.parent Is Nothing Then
        '...
    Else
        ret = lvz.getScreenObject(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
    End If
    getScreenObjectFromNode = ret
End Function

Private Function getImageDefinitionFromMapObjNode(Node As Node) As LVZImageDefinition
    Dim ret As LVZImageDefinition
    Call lvz.setImageDefinitionEmpty(ret)
    If Node Is Nothing Then
        '...
    ElseIf Node.parent Is Nothing Then
        '...
    Else
        Dim tmpMapObj As LVZMapObject
        tmpMapObj = getMapObjectFromNode(Node)
        If tmpMapObj.imgidx <> -1 Then
            ret = lvz.getImageDefinition(getIndexOfLVZFromNode(Node), tmpMapObj.imgidx)
        End If
    End If
    getImageDefinitionFromMapObjNode = ret
End Function

Private Function getImageDefinitionFromScrObjNode(Node As Node) As LVZImageDefinition
    Dim ret As LVZImageDefinition
    Call lvz.setImageDefinitionEmpty(ret)
    If Node Is Nothing Then
        '...
    ElseIf Node.parent Is Nothing Then
        '...
    Else
        Dim tmpScrObj As LVZScreenObject
        tmpScrObj = getScreenObjectFromNode(Node)
        If tmpScrObj.imgidx <> -1 Then
            ret = lvz.getImageDefinition(getIndexOfLVZFromNode(Node), tmpScrObj.imgidx)
        End If
    End If
    getImageDefinitionFromScrObjNode = ret
End Function



Private Function getImageDefinitionFromNode(Node As Node) As LVZImageDefinition
    Dim ret As LVZImageDefinition
    Call lvz.setImageDefinitionEmpty(ret)
    If Node Is Nothing Then
        '...
    ElseIf Node.parent Is Nothing Then
        '...
    Else
        ret = lvz.getImageDefinition(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
    End If
    getImageDefinitionFromNode = ret
End Function

Private Function getIndexOfLVZFromNode(Node As Node) As Integer
    If Node Is Nothing Then
        getIndexOfLVZFromNode = -1
    ElseIf Node.parent Is Nothing Then
        getIndexOfLVZFromNode = -1
    Else
        getIndexOfLVZFromNode = lvz.getIndexOfLVZ(Node.parent.Text)
    End If
End Function

Private Sub tvAvailableFiles_DblClick()

    Dim Node As Node
    Set Node = tvImageDefinitions.SelectedItem
    If Node Is Nothing Then Exit Sub
    If Node.parent Is Nothing Then Exit Sub
    
    Dim imgNode As Node
    Set imgNode = tvAvailableFiles.SelectedItem
    If imgNode Is Nothing Then
        Exit Sub
    ElseIf imgNode.parent Is Nothing Then
        Exit Sub
    End If

    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    Call lvz.setImageDefinitionFile(idx, ParseID(Node.Key), imgNode.Text)
    
    Call UpdateImageDefinitionListData
    
        
    Node.Text = "Image" & ParseID(Node.Key) & IIf(imgNode.Text <> "", " (" & imgNode.Text & ")", " - No File")
    
End Sub

Private Sub tvLVZfiles_Expand(ByVal Node As MSComctlLib.Node)
    Node.selected = True
End Sub

Private Sub tvLVZfiles_NodeClick(ByVal Node As MSComctlLib.Node)
    Call tvLVZfiles_Click
End Sub

Private Sub tvLVZfiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, j As Integer
    Dim path As String
    Dim Node As Node
    Dim imported As Boolean
    imported = False
    
    For j = 1 To Data.files.count
        path = Data.files(j)
        If GetExtension(path) = "lvz" Then
            If lvz.getIndexOfLVZ(GetFileTitle(path)) <> -1 Then
                MessageBox "Could not add lvz file, another file with the same name already exists.", vbOKOnly + vbExclamation
            Else
                Call lvz.importLVZ(path, False)
                imported = True
                
                Set Node = tvLVZfiles.Nodes.add(, , GetFileTitle(path), GetFileTitle(path), 3)
            End If
        Else
            Set Node = tvLVZfiles.HitTest(X, Y)
            If Node Is Nothing Then
            
                If tvLVZfiles.Nodes.count = 0 Then
                    'add new lvz
                    Dim str As String
                    i = lvz.getLVZCount
                    
                    str = "Lvz" & i
                    
                    While lvz.getIndexOfLVZ(str) <> -1
                        i = i + 1
                        str = "Lvz" & i
                    Wend
                    
                    If lvz.AddLVZ(str) Then
                        Set Node = tvLVZfiles.Nodes.add(, , str, str, 3)
                    End If
                End If
                
                If lvz.getLVZCount > 0 Then
                    Call AddFile(path, lvz.getLVZCount - 1)
                End If
                
            Else
                If Node.parent Is Nothing Then
                    Call AddFile(path, lvz.getIndexOfLVZ(Node.Text))
                Else
                    Call AddFile(path, lvz.getIndexOfLVZ(Node.parent.Text))
                End If
            End If
            
            
        
        End If
    Next
    
    If Node Is Nothing Then
    
    Else
        Node.selected = True
        Node.Expanded = True
    End If
    If imported Then
        Call RebuildFilesTree
    End If
    Call UpdateEnabledStatus
End Sub

Private Sub tvLVZfiles_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim Node As Node
    Set Node = tvLVZfiles.HitTest(X, Y)
    If Node Is Nothing Then Exit Sub
    Node.selected = True
    Node.Expanded = True
End Sub

Private Sub tvMapObjects_Click()
    Call RefreshAvailableMapImages
    Call UpdateMapObjectListData
    Call UpdateEnabledStatus

    Call showpreview(getImageDefinitionFromMapObjNode(tvMapObjects.SelectedItem))
    
End Sub

'Private Function isImageDependedOn(lvzidx As Integer, imgidx As Integer) As Boolean
'    Dim i As Integer
'    For i = 0 To lvz.getMapobjectCount(lvzidx) - 1
'        If lvz.getMapObject(lvzidx, i).imgidx = imgidx Then
'            isImageDependedOn = True
'            Exit Function
'        End If
'    Next
'
'    For i = 0 To lvz.getScreenobjectCount(lvzidx) - 1
'        If lvz.getScreenObject(lvzidx, i).imgidx = imgidx Then
'            isImageDependedOn = True
'            Exit Function
'        End If
'    Next
'
'    isImageDependedOn = False
'End Function

'Private Function isFileDependedOn(FileName As String) As Boolean
'    Dim i As Integer
'    Dim j As Integer
'
'    For j = 0 To lvz.getLVZCount - 1
'        For i = 0 To lvz.getImageDefinitionCount(j) - 1
'            If lvz.getImageDefinition(j, i).imagename = FileName Then
'                isFileDependedOn = True
'                Exit Function
'            End If
'        Next
'    Next
'
'    isFileDependedOn = False
'End Function



Sub UpdateEnabledStatus()
    If tvLVZfiles.SelectedItem Is Nothing Then
        cmdRemoveItem.Enabled = False
        cmdRemoveLVZ.Enabled = False
        cmdExport.Enabled = False
        cmdAddItem.Enabled = False
        cmdAddLVZ.Enabled = True
    ElseIf tvLVZfiles.SelectedItem.parent Is Nothing Then
        cmdRemoveItem.Enabled = False
        cmdRemoveLVZ.Enabled = True
        cmdExport.Enabled = True
        cmdAddItem.Enabled = True
        cmdAddLVZ.Enabled = True
    ElseIf Not tvLVZfiles.SelectedItem.parent Is Nothing Then
        cmdRemoveItem.Enabled = True
        cmdRemoveLVZ.Enabled = True
        cmdExport.Enabled = True
        cmdAddItem.Enabled = True
        cmdAddLVZ.Enabled = True
    End If

    If tvMapObjects.SelectedItem Is Nothing Then
        cmdRemoveMapObject.Enabled = False
        cmdAddMapObject.Enabled = False
        lstMapObjectProperties.Enabled = False
    ElseIf tvMapObjects.SelectedItem.parent Is Nothing Then
        cmdRemoveMapObject.Enabled = False
        cmdAddMapObject.Enabled = True
        lstMapObjectProperties.Enabled = False
    ElseIf Not tvMapObjects.SelectedItem.parent Is Nothing Then
        cmdRemoveMapObject.Enabled = True
        cmdAddMapObject.Enabled = True
        lstMapObjectProperties.Enabled = True
        Call UpdateMapObjectListData
    End If
    
    If tvScreenObjects.SelectedItem Is Nothing Then
        cmdRemoveScreenObject.Enabled = False
        cmdAddScreenObject.Enabled = False
        lstScreenObjectProperties.Enabled = False
    ElseIf tvScreenObjects.SelectedItem.parent Is Nothing Then
        cmdRemoveScreenObject.Enabled = False
        cmdAddScreenObject.Enabled = True
        lstScreenObjectProperties.Enabled = False
    ElseIf Not tvScreenObjects.SelectedItem.parent Is Nothing Then
        cmdRemoveScreenObject.Enabled = True
        cmdAddScreenObject.Enabled = True
        lstScreenObjectProperties.Enabled = True
        Call UpdateScreenObjectListData
    End If

    If tvImageDefinitions.SelectedItem Is Nothing Then
        cmdRemoveImgDef.Enabled = False
        cmdAddImgDef.Enabled = False
        lstImageDefinitionProperties.Enabled = False
    ElseIf tvImageDefinitions.SelectedItem.parent Is Nothing Then
        cmdRemoveImgDef.Enabled = False
        cmdAddImgDef.Enabled = True
        lstImageDefinitionProperties.Enabled = False
    ElseIf Not tvImageDefinitions.SelectedItem.parent Is Nothing Then
        cmdRemoveImgDef.Enabled = True
        cmdAddImgDef.Enabled = True
        lstImageDefinitionProperties.Enabled = True
        Call UpdateImageDefinitionListData
    End If
    
End Sub

Sub UpdateMapObjectListData()
    'lstMapObjectProperties.ListItems.Clear

    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    'only show stuff if we have actually selected a map object
    If Node Is Nothing Then
        Exit Sub
    ElseIf Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    objidx = ParseID(Node.Key)
    
    If objidx = -1 Then Exit Sub
    
    Dim tmpMapObj As LVZMapObject
    tmpMapObj = lvz.getMapObject(idx, objidx)
    
    Call lstMapObjectProperties.setPropertyValue(LABEL_x, CStr(tmpMapObj.X))
    Call lstMapObjectProperties.setPropertyValue(LABEL_y, CStr(tmpMapObj.Y))
    If val(tmpMapObj.imgidx) = -1 Then
        Call lstMapObjectProperties.setPropertyValue(LABEL_image, "")
    Else
        Call lstMapObjectProperties.setPropertyValue(LABEL_image, "Image" & tmpMapObj.imgidx & " (" & lvz.getImageDefinition(idx, tmpMapObj.imgidx).imagename & ")")
    End If
    Call lstMapObjectProperties.setPropertyValue(LABEL_layer, CStr(tmpMapObj.layer))
    Call lstMapObjectProperties.setPropertyValue(LABEL_mode, CStr(tmpMapObj.mode))
    
    Call lstMapObjectProperties.setPropertyValue(LABEL_displayTime, CStr(tmpMapObj.displayTime))
    Call lstMapObjectProperties.setPropertyToolTipText(LABEL_displayTime, "Time that the image should be displayed in 1/10th of seconds. Has no effect if Display Mode is set to 'Show Always'")
    
    Call lstMapObjectProperties.setPropertyValue(LABEL_objectID, CStr(tmpMapObj.objectID))
    Call lstMapObjectProperties.setPropertyToolTipText(LABEL_objectID, "*objon #")
End Sub

Sub UpdateScreenObjectListData()
    'lstMapObjectProperties.ListItems.Clear

    Dim Node As Node
    Set Node = tvScreenObjects.SelectedItem
    'only show stuff if we have actually selected a map object
    If Node Is Nothing Then
        Exit Sub
    ElseIf Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    objidx = ParseID(Node.Key)
    
    If objidx = -1 Then Exit Sub
    
    Dim tmpscreenObj As LVZScreenObject
    tmpscreenObj = lvz.getScreenObject(idx, objidx)
    
    Call lstScreenObjectProperties.setPropertyValue(LABEL_x, CStr(tmpscreenObj.X))
    Call lstScreenObjectProperties.setPropertyValue(LABEL_y, CStr(tmpscreenObj.Y))
    Call lstScreenObjectProperties.setPropertyValue(LABEL_typeX, CStr(tmpscreenObj.typeX))
    Call lstScreenObjectProperties.setPropertyValue(LABEL_typeY, CStr(tmpscreenObj.typeY))
    If val(tmpscreenObj.imgidx) = -1 Then
        Call lstScreenObjectProperties.setPropertyValue(LABEL_image, "")
    Else
        Call lstScreenObjectProperties.setPropertyValue(LABEL_image, "Image" & tmpscreenObj.imgidx & " (" & lvz.getImageDefinition(idx, tmpscreenObj.imgidx).imagename & ")")
    End If
    Call lstScreenObjectProperties.setPropertyValue(LABEL_layer, CStr(tmpscreenObj.layer))
    Call lstScreenObjectProperties.setPropertyValue(LABEL_mode, CStr(tmpscreenObj.mode))
    Call lstScreenObjectProperties.setPropertyValue(LABEL_displayTime, CStr(tmpscreenObj.displayTime))
    Call lstScreenObjectProperties.setPropertyValue(LABEL_objectID, CStr(tmpscreenObj.objectID))
End Sub

Sub UpdateImageDefinitionListData()
    'lstMapObjectProperties.ListItems.Clear

    Dim Node As Node
    Set Node = tvImageDefinitions.SelectedItem
    'only show stuff if we have actually selected a map object
    If Node Is Nothing Then
        Exit Sub
    ElseIf Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Integer
    objidx = ParseID(Node.Key)
    
    If objidx = -1 Then Exit Sub
    
    Dim tmpImgDef As LVZImageDefinition
    tmpImgDef = lvz.getImageDefinition(idx, objidx)
    
    Call lstImageDefinitionProperties.setPropertyValue(LABEL_imagename, CStr(tmpImgDef.imagename))
    Call lstImageDefinitionProperties.setPropertyValue(LABEL_animationFramesX, CStr(tmpImgDef.animationFramesX))
    Call lstImageDefinitionProperties.setPropertyValue(LABEL_animationFramesY, CStr(tmpImgDef.animationFramesY))
    Call lstImageDefinitionProperties.setPropertyValue(LABEL_animationTime, CStr(tmpImgDef.animationTime))
End Sub

Private Sub tvMapObjects_DblClick()
    Dim Node As Node
    Set Node = tvMapObjects.SelectedItem
    'only show stuff if we have actually selected a map object
    If Node Is Nothing Then
        Exit Sub
    ElseIf Node.parent Is Nothing Then
        Exit Sub
    End If
    
    Dim idx As Integer
    idx = lvz.getIndexOfLVZ(Node.parent.Text)
    
    If idx = -1 Then
        'parent LVZ not found
        Exit Sub
    End If
    
    
    Dim objidx As Long
    objidx = ParseID(Node.Key)
    
    If objidx = -1 Then Exit Sub
    
    Dim tmpMapObj As LVZMapObject
    tmpMapObj = lvz.getMapObject(idx, objidx)
    
    
    Call parent.SetFocusAt(tmpMapObj.X \ 16, tmpMapObj.Y \ 16, parent.picPreview.width \ 2, parent.picPreview.height \ 2, True)

End Sub

Private Sub tvMapObjects_NodeClick(ByVal Node As MSComctlLib.Node)
    Call tvMapObjects_Click
End Sub

Private Sub tvScreenObjects_Click()
    Call RefreshAvailableScreenImages
    Call UpdateScreenObjectListData
    Call UpdateEnabledStatus
    
    Call showpreview(getImageDefinitionFromScrObjNode(tvScreenObjects.SelectedItem))
End Sub

Private Sub ShowPreviewFile(path As String, Optional FramesX As Integer = 1, Optional FramesY As Integer = 1, Optional animTime As Integer = 100)
    Const BORDER_LEFT = 8
    Const BORDER_TOP = 16
    Const BORDER_RIGHT = 8
    Const BORDER_BOTTOM = 8
    
    If Not FileExists(path) Then Exit Sub
    
    TimerAnim.Enabled = False
    
    Call ClearPreview
    
    previewFramesX = FramesX
    previewFramesY = FramesY
    
    picPreview.AutoSize = True
    Call LoadPic(picPreview, path)
    picPreview.Refresh
    picPreview.AutoSize = False

    Dim picidx As Integer
    picidx = tblvz.SelectedItem.Index - 1
    
    Dim trueWidth As Integer
    Dim trueHeight As Integer
    trueWidth = picPreview.width \ previewFramesX
    trueHeight = picPreview.height \ previewFramesY
    
    Dim maxWidth As Integer
    Dim maxHeight As Integer
    maxWidth = picPreviewAnim(picidx).Container.width - BORDER_LEFT - BORDER_RIGHT
    maxHeight = picPreviewAnim(picidx).Container.height - BORDER_TOP - BORDER_BOTTOM
    
    Dim newWidth As Integer
    Dim newHeight As Integer
        
    If trueWidth > maxWidth Or _
       trueHeight > maxHeight Then
        Dim ratio As Double
        ratio = doubleMinimum(maxWidth / trueWidth, maxHeight / trueHeight)
        newWidth = Int(trueWidth * ratio)
        newHeight = Int(trueHeight * ratio)
    Else
        'center it
        newWidth = trueWidth
        newHeight = trueHeight
    End If
        
    picPreviewAnim(picidx).width = newWidth * Screen.TwipsPerPixelX
    picPreviewAnim(picidx).height = newHeight * Screen.TwipsPerPixelY
    picPreviewAnim(picidx).Left = (BORDER_LEFT + (maxWidth - newWidth) \ 2) * Screen.TwipsPerPixelX
    picPreviewAnim(picidx).Top = (BORDER_TOP + (maxHeight - newHeight) \ 2) * Screen.TwipsPerPixelY
    picPreviewAnim(picidx).visible = True

    TimerAnim.Interval = (animTime / (FramesX * FramesY)) * 10
    
    previewFileName = GetFileTitle(path)
    
    Call RefreshPreview(True, True)
    
    TimerAnim.Enabled = True
End Sub

Private Sub showpreview(imgdef As LVZImageDefinition)

    Dim path As String
    path = lvz.SearchFile(imgdef.imagename)
    
    If Not FileExists(path) Then Exit Sub
    
    Call ShowPreviewFile(path, imgdef.animationFramesX, imgdef.animationFramesY, imgdef.animationTime)
    
End Sub

Private Sub ClearPreview()
    Dim i As Integer
    For i = 0 To 3
        picPreviewAnim(i).Cls
        picPreviewAnim(i).visible = False
    Next
End Sub

Private Sub RefreshPreview(resetFrame As Boolean, changeFrame As Boolean)

    
    Static frameID As Long
    
    If resetFrame Then
        frameID = 0
    End If
    
    Dim frameX As Integer
    Dim frameY As Integer
    frameX = (frameID Mod previewFramesX)
    frameY = ((frameID \ previewFramesX) Mod previewFramesY)

    Dim picidx As Integer
    picidx = tblvz.SelectedItem.Index - 1
    
    Dim trueWidth As Integer
    Dim trueHeight As Integer
    trueWidth = picPreview.width \ previewFramesX
    trueHeight = picPreview.height \ previewFramesY
        
    fPreview(picidx).Caption = "Preview: " & previewFileName & " - " & trueWidth & Chr(215) & trueHeight & " pixels - Frame " & (previewFramesX * frameY + frameX) + 1 & "/" & previewFramesX * previewFramesY
                    
    If trueWidth <> picPreviewAnim(picidx).ScaleWidth Or _
       trueHeight > picPreviewAnim(picidx).ScaleHeight Then
        'stretchblt
        SetStretchBltMode picPreviewAnim(picidx).hDC, HALFTONE
        Call StretchBlt(picPreviewAnim(picidx).hDC, 0, 0, picPreviewAnim(picidx).ScaleWidth, picPreviewAnim(picidx).ScaleHeight, picPreview.hDC, trueWidth * frameX, trueHeight * frameY, trueWidth, trueHeight, vbSrcCopy)
    Else
        'bitblt
        Call BitBlt(picPreviewAnim(picidx).hDC, 0, 0, trueWidth, trueHeight, picPreview.hDC, trueWidth * frameX, trueHeight * frameY, vbSrcCopy)
    End If

    picPreviewAnim(picidx).Refresh
    
    If changeFrame Then frameID = (frameID + 1) Mod 1073676289 '32767 * 32767, max # of frames possible

End Sub

Private Sub tvScreenObjects_NodeClick(ByVal Node As MSComctlLib.Node)
    Call tvScreenObjects_Click
End Sub
