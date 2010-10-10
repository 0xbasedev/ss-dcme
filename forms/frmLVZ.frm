VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
         _ExtentX        =   9128
         _ExtentY        =   5530
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
         _ExtentX        =   9128
         _ExtentY        =   5530
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
         _ExtentX        =   6376
         _ExtentY        =   5318
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
10        Set parent = frm
20        Call lvz.setParent(frm)
End Sub

Private Sub cmdAddImgDef_Click()
          Dim val As Integer
          
          Dim Node As Node
10        Set Node = tvImageDefinitions.SelectedItem
          
20        If Node Is Nothing Then Exit Sub
          
          'make sure we always have the lvz node (and not one of the files)
30        If Not Node.parent Is Nothing Then
40            Set Node = Node.parent
50        End If
                  
          'we selected an lvz
          Dim lvzidx As Integer
60        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
70        If lvzidx <> -1 Then
              Dim obj As LVZImageDefinition
80            Call lvz.setImageDefinitionEmpty(obj)
90            val = lvz.AddImageDefinitionToLVZ(lvzidx, obj)
              
              Dim nd As Node
              'Note: The actual ID of the Image
100           Set nd = tvImageDefinitions.Nodes.add(Node.Text, tvwChild, lvz.getLVZname(lvzidx) & "_Image_" & val, "Image" & val & " - No File")
110           nd.selected = True
120           nd.parent.Expanded = True
              
130           UpdateEnabledStatus
140       End If
End Sub

Private Sub cmdAddItem_Click()
10        On Error GoTo errh
          
          Dim i As Integer
          
          Dim Node As Node
20        Set Node = tvLVZfiles.SelectedItem
          
30        If Node Is Nothing Then Exit Sub
          
          'make sure we always have the lvz node (and not one of the files)
40        If Not Node.parent Is Nothing Then
50            Set Node = Node.parent
60        End If
                  
          'we selected an lvz
          Dim lvzidx As Integer
70        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
80        If lvzidx <> -1 Then
90            cd.DialogTitle = "Add file"
100           cd.flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer
110           cd.Filter = "*.*|*.*"
              
120           cd.ShowOpen
              
              Dim paths() As String
130           paths = ExtractFilePaths(cd.filename)

140           For i = 1 To UBound(paths)
150               Call AddFile(paths(i), lvzidx)
160           Next
              
170       End If
          
180   Exit Sub
errh:
190       If Err = cdlCancel Then
200           Exit Sub
210       End If
End Sub

Private Function MakeKeyForMapObject(lvzidx As Integer, objidx As Long)
10        MakeKeyForMapObject = lvz.getLVZname(lvzidx) & "_MapObject_" & objidx
End Function


Private Sub AddFile(path As String, lvzidx As Integer, Optional showpreview As Boolean = True)
10        If FileExists(path) Then
20            If lvz.SearchFile(path) <> "" Then
30                MessageBox "There is already a file named " & GetFileTitle(path) & " in one of your lvz files.", vbExclamation + vbOKOnly, "Add file"
40            ElseIf lvz.AddFileToLVZ(lvzidx, path) <> -1 Then
                  Dim Node As Node
50                Set Node = tvLVZfiles.Nodes.add(lvz.getLVZname(lvzidx), tvwChild, lvz.getLVZname(lvzidx) & "_" & GetFileTitle(path) & "_" & lvz.getFileCount(lvzidx) - 1, GetFileTitle(path), CInt(lvz.getLVZFileType(lvz.getFileData(lvzidx, lvz.getFileCount(lvzidx) - 1).path)))
60                Node.selected = True
70                Node.parent.Expanded = True
                              
80                Call UpdateEnabledStatus
90                If lvz.getLVZFileType(path) = lvz_image Then
                      Dim previewimg As LVZImageDefinition
100                   previewimg.imagename = GetFileTitle(path)
                      
110                   If Not SetDefaultAnimationProperties(previewimg) Then
                          'Add image definition
120                       If lvz.AddImageDefinitionToLVZ(lvzidx, previewimg) = -1 Then
                              'Could not add (>256)
                              
130                       End If
140                   End If
                      
150                   If showpreview Then Call ShowPreviewFile(path, previewimg.animationFramesX, previewimg.animationFramesY, previewimg.animationTime)
                      
                      
                      
160               End If
170           Else
180               MessageBox "Error adding file " & GetFileTitle(path), vbExclamation + vbOKOnly, "Add file"
190           End If
200       Else
210           MessageBox "File " & path & " not found", vbOKOnly + vbExclamation
220       End If
End Sub
Private Sub cmdAddLVZ_Click()
          Dim str As String
10        str = InputBox("New LVZ Name ?", , "Lvz" & lvz.getLVZCount)
          
20        If str = "" Then
30            Exit Sub
40        End If
          
50        If GetExtension(str) <> "lvz" Then str = str & ".lvz"
          
60        If lvz.AddLVZ(str) Then
              Dim nd As Node
70            Set nd = tvLVZfiles.Nodes.add(, , str, str, 3)
80            nd.selected = True
90            nd.Expanded = True
100       Else
110           MessageBox "Could not add lvz file, another file with the same name already exists.", vbOKOnly + vbExclamation
120       End If
          
130       Call UpdateEnabledStatus
End Sub

Private Sub cmdAddMapObject_Click()
          Dim val As Integer
          
          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
          
20        If Node Is Nothing Then Exit Sub
          
          'make sure we always have the lvz node (and not one of the files)
30        If Not Node.parent Is Nothing Then
40            Set Node = Node.parent
50        End If
                  
          'we selected an lvz
          Dim lvzidx As Integer
60        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
70        If lvzidx <> -1 Then
              Dim obj As LVZMapObject
80            Call lvz.setMapObjectEmpty(obj)
90            val = lvz.AddMapObjectToLVZ(lvzidx, obj)
              
              Dim nd As Node
100           Set nd = tvMapObjects.Nodes.add(Node.Text, tvwChild, lvz.getLVZname(lvzidx) & "_MapObject_" & val, "MapObject" & val)
110           nd.selected = True
120           nd.parent.Expanded = True
              
130           UpdateEnabledStatus
140       End If
End Sub

Private Sub cmdAddScreenObject_Click()
          Dim val As Integer
          
          Dim Node As Node
10        Set Node = tvScreenObjects.SelectedItem
          
20        If Node Is Nothing Then Exit Sub
          
          'make sure we always have the lvz node (and not one of the files)
30        If Not Node.parent Is Nothing Then
40            Set Node = Node.parent
50        End If
                  
          'we selected an lvz
          Dim lvzidx As Integer
60        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
70        If lvzidx <> -1 Then
              Dim obj As LVZScreenObject
80            Call lvz.setScreenObjectEmpty(obj)
90            val = lvz.AddScreenObjectToLVZ(lvzidx, obj)
              
              Dim nd As Node
100           Set nd = tvScreenObjects.Nodes.add(Node.Text, tvwChild, lvz.getLVZname(lvzidx) & "_ScreenObject_" & val, "ScreenObject" & val)
110           nd.selected = True
120           nd.parent.Expanded = True
              
130           UpdateEnabledStatus
140       End If
End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub cmdExport_Click()
10        On Error GoTo errh
          
          Dim Node As Node
20        Set Node = tvLVZfiles.SelectedItem
30        If Not Node.parent Is Nothing Then
40            Set Node = Node.parent
50        End If
          
          Dim lvzidx As Integer
60        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
70        If lvzidx = -1 Then
              'lvz not found
80            Exit Sub
90        End If
          
          
100       cd.DialogTitle = "Export LVZ"
110       cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
120       cd.Filter = "*.lvz|*.lvz"
          
130       cd.ShowSave
          
140       Call lvz.exportLVZ(cd.filename, lvzidx)
          
150       Call UpdateEnabledStatus
          
160   Exit Sub
errh:
170       If Err = cdlCancel Then
180           Exit Sub
190       End If

End Sub

Private Sub cmdImport_Click()
10        On Error GoTo errh
          
20        cd.DialogTitle = "Import LVZ"
30        cd.flags = cdlOFNHideReadOnly
40        cd.Filter = "*.lvz|*.lvz"
          
50        cd.ShowOpen
          
60        Call lvz.importLVZ(cd.filename, False)
          
70        Call RebuildFilesTree
80        Call UpdateMapObjectListData
90        Call RebuildMapObjectsTree
100       Call UpdateEnabledStatus
          
110   Exit Sub
errh:
120       If Err = cdlCancel Then
130           Exit Sub
140       End If
End Sub

Private Sub cmdMakeIni_Click()
10        On Error GoTo errh
          
          Dim Node As Node
20        Set Node = tvLVZfiles.SelectedItem
30        If Not Node.parent Is Nothing Then
40            Set Node = Node.parent
50        End If
          
          Dim lvzidx As Integer
60        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
70        If lvzidx = -1 Then
              'lvz not found
80            Exit Sub
90        End If
          
          
100       cd.DialogTitle = "Export INI"
110       cd.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
120       cd.Filter = "*.ini|*.ini"
          
130       cd.ShowSave
          
140       Call lvz.exportINI(cd.filename, lvzidx)
          
150       Call UpdateEnabledStatus
          
160   Exit Sub
errh:
170       If Err = cdlCancel Then
180           Exit Sub
190       End If
End Sub

Private Sub cmdOK_Click()
          
          'check if each object has an image assigned
          Dim lvzidx As Integer
          Dim obj As Long, Img As Integer
          
10        For lvzidx = 0 To lvz.getLVZCount - 1
20            For Img = 0 To lvz.getImageDefinitionCount(lvzidx) - 1
30                If lvz.getImageDefinition(lvzidx, Img).imagename = "" Then
40                    Call MessageBox("Image" & Img & " from lvz '" & lvz.getLVZname(lvzidx) & "' has no file assigned!", vbExclamation)
50                    tblvz.Tabs("Images").selected = True
60                    tblvz_Click
70                    tvImageDefinitions.Nodes(lvz.getLVZname(lvzidx) & "_Image_" & Img).selected = True
80                    UpdateEnabledStatus
90                    UpdateImageDefinitionListData
100                   Exit Sub
110               End If
120           Next
              
130           For obj = 0 To lvz.getMapobjectCount(lvzidx) - 1
140               If lvz.getMapObjectImageID(lvzidx, obj) = -1 Or lvz.getMapObjectImageID(lvzidx, obj) >= lvz.getImageDefinitionCount(lvzidx) Then
150                   Call MessageBox("MapObject" & obj & " from lvz '" & lvz.getLVZname(lvzidx) & "' has no image assigned!", vbExclamation)
160                   tblvz.Tabs("MapObjects").selected = True
170                   tblvz_Click
180                   tvMapObjects.Nodes(lvz.getLVZname(lvzidx) & "_MapObject_" & obj).selected = True
190                   UpdateEnabledStatus
200                   UpdateMapObjectListData
210                   RefreshAvailableMapImages
220                   Exit Sub
230               End If
240           Next
              
250           For obj = 0 To lvz.getScreenobjectCount(lvzidx) - 1
260               If lvz.getScreenObjectImageID(lvzidx, obj) = -1 Or lvz.getScreenObjectImageID(lvzidx, obj) >= lvz.getImageDefinitionCount(lvzidx) Then
270                   Call MessageBox("ScreenObject" & obj & " from lvz '" & lvz.getLVZname(lvzidx) & "' has no image assigned!", vbExclamation)
280                   tblvz.Tabs("ScreenObjects").selected = True
290                   tblvz_Click
300                   tvScreenObjects.Nodes(lvz.getLVZname(lvzidx) & "_ScreenObject_" & obj).selected = True
310                   UpdateEnabledStatus
320                   UpdateScreenObjectListData
330                   RefreshAvailableScreenImages
340                   Exit Sub
350               End If
360           Next
370       Next
                      
380       Call parent.lvz.setLVZData(lvz.getLVZData, lvz.getLVZCount)
          
390       Call parent.lvz.buildAllLVZImages
          
          'Call parent.lvz.DrawLVZImageInterface(parent.lvz.curImageFrame)
'400       Call parent.tileset.DrawLVZTileset(True) <-- buildAllLVZImages already redraws the lvz tileset
          
          
410       Unload Me
End Sub

Private Sub cmdRemoveImgDef_Click()
          Dim Node As Node
10        Set Node = tvImageDefinitions.SelectedItem
20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          Dim fileIdx As Integer
          'works for filetitle's too!
90        fileIdx = ParseID(Node.Key)

100       If fileIdx = -1 Then
              'file not found in lvz
110       Else
              Dim MapObjCount As Long, ScrObjCount As Long
120           MapObjCount = lvz.CountMapObjectsUsingImage(idx, fileIdx)
130           ScrObjCount = lvz.CountScreenObjectsUsingImage(idx, fileIdx)
              
140           If MapObjCount > 0 Or ScrObjCount > 0 Then
                  Dim ret As VbMsgBoxResult
150               ret = MessageBox("There are " & MapObjCount & " map objects and " & ScrObjCount & " screen objects associated to this image. Deleting it will also delete all these objects. Do you wish to delete it?" & fileIdx & "?", vbExclamation + vbYesNo, "Confirm Delete Image Definition")
                  
160               If ret = vbNo Then
170                   Exit Sub
180               End If
190           End If
              

200           Call lvz.RemoveLinksToImage(idx, fileIdx)
              
210           Call lvz.removeImageDefinitionFromLVZ(idx, fileIdx)
              'if this doesn't work, use rebuildtree
              'Call tvLVZfiles.Nodes.Remove(Node.Index)
220           Call RebuildFilesTree
230           Call RebuildImageDefinitionsTree
              
240           Call UpdateEnabledStatus

250       End If
End Sub

Private Function ParseID(Text As String, Optional Delimiter As String = "_", Optional Index As Integer = -1) As Integer
10        If Text = "" Then
20            ParseID = -1
30            Exit Function
40        End If
          
          Dim tmp() As String
50        tmp = Split(Text, Delimiter)
          
60        If Index = -1 Then Index = UBound(tmp)
          
70        If UBound(tmp) >= Index Then
80            If IsNumeric(tmp(Index)) Then
90                ParseID = CInt(tmp(Index))
100           Else
110               ParseID = -1
120           End If
130       Else
140           ParseID = -1
150       End If
End Function

Private Sub cmdRemoveItem_Click()
          Dim Node As Node
10        Set Node = tvLVZfiles.SelectedItem

20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
              
          
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          Dim fileIdx As Long
          'works for filetitle's too!
90        fileIdx = lvz.getIndexOfFile(idx, Node.Text)
          

          
100       If fileIdx = -1 Then
              'file not found in lvz
110       Else
              Dim imgCount As Long
120           imgCount = lvz.CountImagesUsingFile(Node.Text)
              
130           If imgCount > 0 Then
                  Dim ret As VbMsgBoxResult
140               ret = MessageBox("There are " & imgCount & " image definitions using '" & Node.Text & "'. Deleting it will also remove all these image definitions, and all objects using them. Do you wish to delete it?", vbExclamation + vbYesNo, "Confirm Remove File")
                  
150               If ret = vbNo Then
160                   Exit Sub
170               End If
180           End If
              

190           Call lvz.RemoveLinksToFile(Node.Text)
              
              
200           Call lvz.removeFileFromLVZ(idx, fileIdx)
              'if this doesn't work, use rebuildtree
              
              Dim nodeIdx As Long
210           nodeIdx = Node.Index
              
220           Call RebuildFilesTree
              
230           Call RestoreNodeSelection(tvLVZfiles, nodeIdx)
              
240           Call tvLVZfiles_Click
              
250           Call UpdateEnabledStatus

              

260       End If

End Sub

Private Sub cmdRemoveLVZ_Click()
          Dim Node As Node
10        Set Node = tvLVZfiles.SelectedItem
20        If Not Node.parent Is Nothing Then
30            Set Node = Node.parent
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.Text)
          
60        If idx = -1 Then
              'lvz not found
70        Else
80            If lvz.getFileCount(idx) > 0 Then
90                If MessageBox("Delete " & lvz.getLVZname(idx) & " and all its files, image definitions, map objects and screen objects?", vbYesNo + vbQuestion, "Confirm Delete LVZ") = vbNo Then
100                   Exit Sub
110               End If
120           End If
              
130           Call lvz.removeLVZ(idx)
              
              Dim nodeIdx As Integer
140           nodeIdx = Node.Index
              
150           Call RebuildFilesTree
              
              'Call RestoreNodeSelection(tvLVZfiles, nodeIdx)
              
160           Call tvLVZfiles_Click
              
170           Call UpdateEnabledStatus
180       End If
          
          
End Sub


Private Sub RestoreNodeSelection(ByRef tree As TreeView, nodeIdx As Long)
          Dim nodecount As Long
10        nodecount = tree.Nodes.count
          
20        If nodeIdx > nodecount Then
30            If nodecount > 0 Then
40                tree.Nodes(nodecount).selected = True
50                tree.SelectedItem.EnsureVisible
60            End If
70        ElseIf nodeIdx > 0 Then
80            tree.Nodes(nodeIdx).selected = True
90            tree.SelectedItem.EnsureVisible
100       End If
End Sub

Private Sub cmdRemoveMapObject_Click()
          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          
          Dim objidx As Long
          'works for filetitle's too!
90        objidx = ParseID(Node.Key)
          
          
          Dim nodeIdx As Long
          'Node.Key
100       nodeIdx = Node.Index
          
110       Call lvz.removeMapObjectFromLVZ(idx, objidx)
          'if this doesn't work, use rebuildtree
          'Call tvMapObjects.Nodes.Remove(Node.Index)
120       Call RebuildMapObjectsTree
          
130       Call RestoreNodeSelection(tvMapObjects, nodeIdx)
          
140       Call tvMapObjects_Click
          
          'tvMapObjects.In
End Sub

Private Sub cmdRemoveScreenObject_Click()
          Dim Node As Node
10        Set Node = tvScreenObjects.SelectedItem
20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          
          Dim objidx As Long
          'works for filetitle's too!
90        objidx = ParseID(Node.Key)
          
          Dim nodeIdx As Long
100       nodeIdx = Node.Index
          
110       Call lvz.removeScreenObjectFromLVZ(idx, objidx)
          
120       Call RebuildScreenObjectsTree
          
130       Call RestoreNodeSelection(tvScreenObjects, nodeIdx)
          
140       Call tvScreenObjects_Click
          
End Sub

Private Sub Form_Load()
         
10        Set Me.Icon = frmGeneral.Icon
          
20        Call lvz.setLVZData(parent.lvz.getLVZData, parent.lvz.getLVZCount)
          
30        Call RebuildFilesTree
40        Call RebuildMapObjectsTree
50        Call RebuildScreenObjectsTree
60        Call RebuildImageDefinitionsTree

70        CreateImageDefinitionsProperties
80        CreateScreenObjectProperties
90        CreateMapObjectProperties
          
          'make sure controls are enabled correctly
100       Call UpdateEnabledStatus
          
110       tblvz_Click
          
End Sub

Private Sub CreateImageDefinitionsProperties()
10        Call lstImageDefinitionProperties.AddProperty(LABEL_imagename, p_text)
20            Call lstImageDefinitionProperties.setPropertyLocked(LABEL_imagename, True)
30        Call lstImageDefinitionProperties.AddProperty(LABEL_animationFramesX, p_number, 1, 32767)
40        Call lstImageDefinitionProperties.AddProperty(LABEL_animationFramesY, p_number, 1, 32767)
50        Call lstImageDefinitionProperties.AddProperty(LABEL_animationTime, p_number, 1, 65535)
End Sub

Private Sub CreateScreenObjectProperties()
          Dim chlst() As String
          
10        Call lstScreenObjectProperties.AddProperty(LABEL_x, p_number, -2048, 2047)
          
20        Call lstScreenObjectProperties.AddProperty(LABEL_typeX, p_list)
30            ReDim chlst(11)
40            chlst(0) = "Screen Left"
50            chlst(1) = "Screen Center"
60            chlst(2) = "Screen Right"
70            chlst(3) = "Stats Box Right Edge"
80            chlst(4) = "Specials Right"
90            chlst(5) = "Specials Right"
100           chlst(6) = "Energy Bar Center"
110           chlst(7) = "Chat Text Right Edge"
120           chlst(8) = "Radar Left Edge"
130           chlst(9) = "Clock Left Edge"
140           chlst(10) = "Weapons Left"
150           chlst(11) = ""
160           Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_typeX, chlst)
              
170       Call lstScreenObjectProperties.AddProperty(LABEL_y, p_number, -2048, 2047)
180           Call lstScreenObjectProperties.AddProperty(LABEL_typeY, p_list)
190           ReDim chlst(11)
200           chlst(0) = "Screen Top"
210           chlst(1) = "Screen Center"
220           chlst(2) = "Screen Bottom"
230           chlst(3) = "Stats Box Bottom Edge"
240           chlst(4) = "Specials Top"
250           chlst(5) = "Specials Bottom"
260           chlst(6) = "Under Energy Bar"
270           chlst(7) = "Chat Text Top"
280           chlst(8) = "Radar Top"
290           chlst(9) = "Clock Top"
300           chlst(10) = "Weapons Top"
310           chlst(11) = "Weapons Bottom"
320           Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_typeY, chlst)
              
330       Call lstScreenObjectProperties.AddProperty(LABEL_image, p_text)
340           Call lstScreenObjectProperties.setPropertyLocked(LABEL_image, True)
350       Call lstScreenObjectProperties.AddProperty(LABEL_layer, p_list)

360           ReDim chlst(7)
370           chlst(0) = "0 - Below All"
380           chlst(1) = "1 - After Background"
390           chlst(2) = "2 - After Tiles"
400           chlst(3) = "3 - After Weapons"
410           chlst(4) = "4 - After Ships"
420           chlst(5) = "5 - After Gauges"
430           chlst(6) = "6 - After Chat"
440           chlst(7) = "7 - Top Most"
450           Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_layer, chlst)
460       Call lstScreenObjectProperties.AddProperty(LABEL_mode, p_list)
470           ReDim chlst(5)
480           chlst(0) = "0 - Show Always"
490           chlst(1) = "1 - Enter Zone"
500           chlst(2) = "2 - Enter Arena"
510           chlst(3) = "3 - Kill"
520           chlst(4) = "4 - Death"
530           chlst(5) = "5 - Server Controlled"
540           Call lstScreenObjectProperties.setPropertyChoiceList(LABEL_mode, chlst)
550       Call lstScreenObjectProperties.AddProperty(LABEL_displayTime, p_number)
560           Call lstScreenObjectProperties.setPropertyNumberBoundaries(LABEL_displayTime, 0, 100000000)
          
570       Call lstScreenObjectProperties.AddProperty(LABEL_objectID, p_number)
580           Call lstScreenObjectProperties.setPropertyNumberBoundaries(LABEL_objectID, 0, 32767)
End Sub

Private Sub CreateMapObjectProperties()

10        Call lstMapObjectProperties.AddProperty(LABEL_x, p_number, -32767, 32767)
20        Call lstMapObjectProperties.AddProperty(LABEL_y, p_number, -32767, 32767)
30        Call lstMapObjectProperties.AddProperty(LABEL_image, p_text)
40            Call lstMapObjectProperties.setPropertyLocked(LABEL_image, True)
50        Call lstMapObjectProperties.AddProperty(LABEL_layer, p_list)
              Dim chlst() As String
60            ReDim chlst(7)
70            chlst(0) = "0 - Below All"
80            chlst(1) = "1 - After Background"
90            chlst(2) = "2 - After Tiles"
100           chlst(3) = "3 - After Weapons"
110           chlst(4) = "4 - After Ships"
120           chlst(5) = "5 - After Gauges"
130           chlst(6) = "6 - After Chat"
140           chlst(7) = "7 - Top Most"
150           Call lstMapObjectProperties.setPropertyChoiceList(LABEL_layer, chlst)
160       Call lstMapObjectProperties.AddProperty(LABEL_mode, p_list)
170           ReDim chlst(5)
180           chlst(0) = "0 - Show Always"
190           chlst(1) = "1 - Enter Zone"
200           chlst(2) = "2 - Enter Arena"
210           chlst(3) = "3 - Kill"
220           chlst(4) = "4 - Death"
230           chlst(5) = "5 - Server Controlled"
240           Call lstMapObjectProperties.setPropertyChoiceList(LABEL_mode, chlst)
250       Call lstMapObjectProperties.AddProperty(LABEL_displayTime, p_number)
260           Call lstMapObjectProperties.setPropertyNumberBoundaries(LABEL_displayTime, 0, 100000000)
          
270       Call lstMapObjectProperties.AddProperty(LABEL_objectID, p_number)
280           Call lstMapObjectProperties.setPropertyNumberBoundaries(LABEL_objectID, 0, 32767)
End Sub

Private Sub RebuildFilesTree()
10        tvLVZfiles.Nodes.Clear
          Dim i As Integer
          Dim j As Long
          
20        For i = 0 To lvz.getLVZCount - 1
30            Call tvLVZfiles.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
40            For j = 0 To lvz.getFileCount(i) - 1
50                Call tvLVZfiles.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_" & GetFileTitle(lvz.getFileData(i, j).path) & "_" & j, GetFileTitle(lvz.getFileData(i, j).path), CInt(lvz.getLVZFileType(lvz.getFileData(i, j).path)))
60            Next
70        Next
End Sub

Private Sub RebuildMapObjectsTree()
10        tvMapObjects.Nodes.Clear
          Dim i As Integer
          Dim j As Long
          Dim lvzname As String
          
20        tvMapObjects.visible = False
30        For i = 0 To lvz.getLVZCount - 1
40            lvzname = lvz.getLVZname(i)
              
50            Call tvMapObjects.Nodes.add(, , lvzname, lvzname, 3)
              
60            If lvz.getMapobjectCount(i) > 3000 Then
70                Call tvMapObjects.Nodes.add(lvzname, tvwChild, lvzname & "_MapObject_0", "<Too many objects to display (" & lvz.getMapobjectCount(i) & ")>")
80            Else
90                For j = 0 To lvz.getMapobjectCount(i) - 1
          '            MakeKeyForMapObject(i, j)
          '            tvMapObjects.Sorted
100                   Call tvMapObjects.Nodes.add(lvzname, tvwChild, lvzname & "_MapObject_" & j, "MapObject" & j)
          '            Call tvMapObjects.Nodes.add(lvzname, tvwChild, , "MapObject" & j)
110               Next
120           End If
130       Next
          
140       tvMapObjects.visible = True
End Sub

Private Sub RebuildScreenObjectsTree()
10        tvScreenObjects.Nodes.Clear
          Dim i As Integer
          Dim j As Long
          
20        For i = 0 To lvz.getLVZCount - 1
30            Call tvScreenObjects.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
40            For j = 0 To lvz.getScreenobjectCount(i) - 1
50                Call tvScreenObjects.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_ScreenObject_" & j, "ScreenObject" & j)
60            Next
70        Next
End Sub

Private Sub RebuildImageDefinitionsTree()
10        tvImageDefinitions.Nodes.Clear
          Dim i As Integer
          Dim j As Integer
          
20        For i = 0 To lvz.getLVZCount - 1
30            Call tvImageDefinitions.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
40            For j = 0 To lvz.getImageDefinitionCount(i) - 1
50                Call tvImageDefinitions.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_Image_" & j, "Image" & j & IIf(lvz.getImageDefinition(i, j).imagename <> "", " (" & lvz.getImageDefinition(i, j).imagename & ")", " - No File"))
60            Next
70        Next
End Sub

Private Sub RebuildAvailableFilesTree()
10        tvAvailableFiles.Nodes.Clear
          Dim i As Integer
          Dim j As Long
          
20        For i = 0 To lvz.getLVZCount - 1
              Dim nodeCreated As Boolean
30            nodeCreated = False
40            For j = 0 To lvz.getFileCount(i) - 1
                  Dim tmpFiledata As LVZFileStruct
50                tmpFiledata = lvz.getFileData(i, j)
60                If tmpFiledata.Type = lvz_image Then
70                    If Not nodeCreated Then
80                        Call tvAvailableFiles.Nodes.add(, , lvz.getLVZname(i), lvz.getLVZname(i), 3)
90                        nodeCreated = True
100                   End If
110                   Call tvAvailableFiles.Nodes.add(lvz.getLVZname(i), tvwChild, lvz.getLVZname(i) & "_Image_" & j, GetFileTitle(tmpFiledata.path))
120               End If
130           Next
              
140       Next
End Sub

Private Sub RefreshAvailableMapImages()
10        lstMapImgDefs.Clear
          
20        If tvMapObjects.SelectedItem Is Nothing Then
30            Exit Sub
40        End If
          
          Dim Node As Node
50        Set Node = tvMapObjects.SelectedItem
          'make sure we always have the lvz node (and not one of the files)
60        If Not Node.parent Is Nothing Then
70            Set Node = Node.parent
80        End If
          
          Dim lvzidx As Integer
90        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
          Dim i As Integer
          
100       For i = 0 To lvz.getImageDefinitionCount(lvzidx) - 1
110           Call lstMapImgDefs.addItem("Image" & i & IIf(lvz.getImageDefinition(lvzidx, i).imagename <> "", " (" & lvz.getImageDefinition(lvzidx, i).imagename & ")", " - No File"))
120       Next
End Sub

Private Sub RefreshAvailableScreenImages()
10        lstScreenImgDefs.Clear

20        If tvScreenObjects.SelectedItem Is Nothing Then
30            Exit Sub
40        End If
          
          Dim Node As Node
50        Set Node = tvScreenObjects.SelectedItem
          'make sure we always have the lvz node (and not one of the files)
60        If Not Node.parent Is Nothing Then
70            Set Node = Node.parent
80        End If
          
          Dim lvzidx As Integer
90        lvzidx = lvz.getIndexOfLVZ(Node.Text)
          
          Dim i As Integer
          
100       For i = 0 To lvz.getImageDefinitionCount(lvzidx) - 1
110           Call lstScreenImgDefs.addItem("Image" & i & IIf(lvz.getImageDefinition(lvzidx, i).imagename <> "", " (" & lvz.getImageDefinition(lvzidx, i).imagename & ")", " - No File"))
120       Next
          
End Sub



Private Sub Form_Unload(Cancel As Integer)
10        Call lvz.setParent(Nothing)
20        Set parent = Nothing
          
End Sub

Private Sub lstImageDefinitionProperties_LostFocus()
10        Call lstImageDefinitionProperties.EditOff(True)
End Sub

Private Sub lstImageDefinitionProperties_PropertyChanged(propName As String)
          Dim Node As Node
10        Set Node = tvImageDefinitions.SelectedItem
20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          
          Dim imgidx As Integer
          'works for filetitle's too!
90        imgidx = ParseID(Node.Key)
          
          Dim tmpImgDef As LVZImageDefinition
100       tmpImgDef = lvz.getImageDefinition(idx, imgidx)
          
110       Select Case propName
              Case LABEL_imagename
120               tmpImgDef.imagename = val(lstImageDefinitionProperties.getPropertyValue(propName))
130           Case LABEL_animationFramesX
140               tmpImgDef.animationFramesX = val(lstImageDefinitionProperties.getPropertyValue(propName))
150           Case LABEL_animationFramesY
160               tmpImgDef.animationFramesY = val(lstImageDefinitionProperties.getPropertyValue(propName))
170           Case LABEL_animationTime
180               tmpImgDef.animationTime = val(lstImageDefinitionProperties.getPropertyValue(propName))
190       End Select
          
200       Call lvz.setImageDefinition(idx, imgidx, tmpImgDef)
          
          Dim path As String
210       path = lvz.SearchFile(tmpImgDef.imagename)
          
220       If FileExists(path) Then
230           Call showpreview(tmpImgDef)
240       End If
End Sub

Private Sub lstMapImgDefs_Click()
          Dim imgidx As Integer
          Dim i As Integer
          
          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
20        If Node Is Nothing Then Exit Sub

          Dim idx As Integer
30        If Node.parent Is Nothing Then
40            idx = lvz.getIndexOfLVZ(Node.Text)
50        Else
60            idx = lvz.getIndexOfLVZ(Node.parent.Text)
70        End If
          
80        For i = 0 To lstMapImgDefs.ListCount - 1
90            If lstMapImgDefs.selected(i) Then
100               imgidx = i
110               Exit For
120           End If
130       Next
          
          Dim tmpImg As LVZImageDefinition
140       tmpImg = lvz.getImageDefinition(idx, imgidx)
          
150       Call showpreview(lvz.getImageDefinition(idx, imgidx))
End Sub

Private Sub lstMapImgDefs_DblClick()
          Dim newImgIdx As Integer
          Dim i As Integer

          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
20        If Node Is Nothing Then Exit Sub
30        If Node.parent Is Nothing Then Exit Sub
          
40        For i = 0 To lstMapImgDefs.ListCount - 1
50            If lstMapImgDefs.selected(i) Then
60                newImgIdx = i
70                Exit For
80            End If
90        Next
          

          Dim idx As Integer
100       idx = lvz.getIndexOfLVZ(Node.parent.Text)
              
110       If idx = -1 Then
              'parent LVZ not found
120           Exit Sub
130       End If

          Dim objidx As Long
          'works for filetitle's too!
140       objidx = ParseID(Node.Key)
          
150       Call lvz.setMapObjectImageID(idx, objidx, newImgIdx)
          
160       Call UpdateMapObjectListData
End Sub

Private Sub lstMapObjectProperties_LostFocus()
10        Call lstMapObjectProperties.EditOff(True)
End Sub

Private Sub lstMapObjectProperties_PropertyChanged(propName As String)
          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          
          Dim objidx As Long
          'works for filetitle's too!
90        objidx = ParseID(Node.Key)
          
          Dim tmpMapObj As LVZMapObject
100       tmpMapObj = lvz.getMapObject(idx, objidx)
          
110       Select Case propName
              Case LABEL_x
120               tmpMapObj.X = val(lstMapObjectProperties.getPropertyValue(propName))
130           Case LABEL_y
140               tmpMapObj.Y = val(lstMapObjectProperties.getPropertyValue(propName))
150           Case LABEL_layer
                  Dim newlayer As LVZLayerEnum
                  
160               newlayer = val(lstMapObjectProperties.getPropertyValue(propName))
170               tmpMapObj.layer = newlayer
                  
                  
                  'We have to keep it sorted
180               objidx = lvz.ChangeMapObjectLayer(idx, objidx, newlayer)
                  
190               Call RebuildMapObjectsTree
                  
      '            Dim nodeIdx As Integer
                  
                  '                                       indexes start at 1, and we don't want the parent 'MapObjects' node
      '            Call RestoreNodeSelection(tvMapObjects, objidx + 2)
200               tvMapObjects.Nodes(MakeKeyForMapObject(idx, objidx)).selected = True
210               tvMapObjects.SelectedItem.EnsureVisible
                  
220           Case LABEL_mode
230               tmpMapObj.mode = val(lstMapObjectProperties.getPropertyValue(propName))
240           Case LABEL_displayTime
250               tmpMapObj.displayTime = val(lstMapObjectProperties.getPropertyValue(propName))
260           Case LABEL_objectID
270               tmpMapObj.objectID = val(lstMapObjectProperties.getPropertyValue(propName))
280       End Select
          
290       Call lvz.setMapObject(idx, objidx, tmpMapObj)
End Sub



Private Sub lstScreenImgDefs_Click()
          Dim imgidx As Integer
          Dim i As Integer
          
          Dim Node As Node
10        Set Node = tvScreenObjects.SelectedItem
20        If Node Is Nothing Then Exit Sub
          
          Dim idx As Integer
30        If Node.parent Is Nothing Then
40            idx = lvz.getIndexOfLVZ(Node.Text)
50        Else
60            idx = lvz.getIndexOfLVZ(Node.parent.Text)
70        End If
          
80        For i = 0 To lstScreenImgDefs.ListCount - 1
90            If lstScreenImgDefs.selected(i) Then
100               imgidx = i
110               Exit For
120           End If
130       Next
          

140       Call showpreview(lvz.getImageDefinition(idx, imgidx))
End Sub

Private Sub lstScreenImgDefs_DblClick()
          Dim newImgIdx As Integer
          Dim i As Integer

          Dim Node As Node
10        Set Node = tvScreenObjects.SelectedItem
20        If Node Is Nothing Then Exit Sub
30        If Node.parent Is Nothing Then Exit Sub
          
40        For i = 0 To lstScreenImgDefs.ListCount - 1
50            If lstScreenImgDefs.selected(i) Then
60                newImgIdx = i
70                Exit For
80            End If
90        Next
          

          Dim idx As Integer
100       idx = lvz.getIndexOfLVZ(Node.parent.Text)
              
110       If idx = -1 Then
              'parent LVZ not found
120           Exit Sub
130       End If

          Dim objidx As Long
          'works for filetitle's too!
140       objidx = ParseID(Node.Key)
          
150       Call lvz.setScreenObjectImageID(idx, objidx, newImgIdx)
          
160       Call UpdateScreenObjectListData
End Sub

Private Sub lstScreenImgDefs_Scroll()
10        Call lstScreenImgDefs_Click
End Sub

Private Sub lstScreenObjectProperties_LostFocus()
10        Call lstScreenObjectProperties.EditOff(True)
End Sub

Private Sub lstScreenObjectProperties_PropertyChanged(propName As String)
          Dim Node As Node
10        Set Node = tvScreenObjects.SelectedItem
20        If Node.parent Is Nothing Then
30            Exit Sub
40        End If
          
          Dim idx As Integer
50        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
60        If idx = -1 Then
              'parent LVZ not found
70            Exit Sub
80        End If
          
          
          Dim objidx As Long
          'works for filetitle's too!
90        objidx = ParseID(Node.Key)
          
          Dim tmpScrObj As LVZScreenObject
100       tmpScrObj = lvz.getScreenObject(idx, objidx)
          
110       Select Case propName
              Case LABEL_x
120               tmpScrObj.X = val(lstScreenObjectProperties.getPropertyValue(propName))
130           Case LABEL_y
140               tmpScrObj.Y = val(lstScreenObjectProperties.getPropertyValue(propName))
150           Case LABEL_layer
160               tmpScrObj.layer = val(lstScreenObjectProperties.getPropertyValue(propName))
170           Case LABEL_mode
180               tmpScrObj.mode = val(lstScreenObjectProperties.getPropertyValue(propName))
190           Case LABEL_displayTime
200               tmpScrObj.displayTime = val(lstScreenObjectProperties.getPropertyValue(propName))
210           Case LABEL_objectID
220               tmpScrObj.objectID = val(lstScreenObjectProperties.getPropertyValue(propName))
230           Case LABEL_typeX
240               tmpScrObj.typeX = val(lstScreenObjectProperties.getPropertyValue(propName))
250           Case LABEL_typeY
260               tmpScrObj.typeY = val(lstScreenObjectProperties.getPropertyValue(propName))
                  
270       End Select
          
280       Call lvz.setScreenObject(idx, objidx, tmpScrObj)
End Sub

Private Sub tblvz_BeforeClick(Cancel As Integer)
10        TimerAnim.Enabled = False
20        Call ClearPreview
          
30        If tblvz.SelectedItem.Key = "MapObjects" Then
40            Call lstMapObjectProperties.EditOff(True)
50        ElseIf tblvz.SelectedItem.Key = "ScreenObjects" Then
60            Call lstScreenObjectProperties.EditOff(True)
70        ElseIf tblvz.SelectedItem.Key = "Images" Then
80            Call lstImageDefinitionProperties.EditOff(True)
90        End If
End Sub

Private Sub tblvz_Click()
          Dim i As Integer
          'show and enable the selected tblvz's controls
          'and hide and disable all others

          
10        For i = 0 To tblvz.Tabs.count - 1
20            If i = tblvz.SelectedItem.Index - 1 Then
30                picTabContents(i).Left = tblvz.Left + 8
40                picTabContents(i).Enabled = True
50                picTabContents(i).visible = True
60                picTabContents(i).Top = tblvz.Top + 24
70                picTabContents(i).width = tblvz.width - 16
80                picTabContents(i).height = tblvz.height - 32
90            Else
100               picTabContents(i).Left = -5000
110               picTabContents(i).Enabled = False
120               picTabContents(i).visible = False
130           End If
140       Next
          
150       If tblvz.SelectedItem.Index = 1 Then
160           Call RebuildFilesTree
              
170       ElseIf tblvz.SelectedItem.Index = 2 Then
180           Call RebuildImageDefinitionsTree
190           Call RebuildAvailableFilesTree
              
200       ElseIf tblvz.SelectedItem.Index = 3 Then
210           Call RebuildMapObjectsTree
220           Call RefreshAvailableMapImages
              
230       ElseIf tblvz.SelectedItem.Index = 4 Then
240           Call RebuildScreenObjectsTree
250           Call RefreshAvailableScreenImages
              
260       End If
270       Call UpdateEnabledStatus
End Sub



Private Sub TimerAnim_Timer()
10        Call RefreshPreview(False, True)
End Sub

Private Sub tvAvailableFiles_Click()
          Dim tmpfile As LVZFileStruct
          
10        tmpfile = getFileFromNode(tvAvailableFiles.SelectedItem)
          
20        If FileExists(tmpfile.path) Then
30            If tmpfile.Type = lvz_image Then
40                Call ShowPreviewFile(tmpfile.path)
50            End If
60        End If
End Sub

Private Sub tvAvailableFiles_NodeClick(ByVal Node As MSComctlLib.Node)
10        Call tvAvailableFiles_Click
End Sub

Private Sub tvImageDefinitions_Click()
10        Call UpdateImageDefinitionListData
20        Call UpdateEnabledStatus
          
30        Call showpreview(getImageDefinitionFromNode(tvImageDefinitions.SelectedItem))

End Sub

Private Sub tvImageDefinitions_NodeClick(ByVal Node As MSComctlLib.Node)
10        Call tvImageDefinitions_Click
End Sub

Private Sub tvLVZfiles_Click()
10        Call UpdateEnabledStatus
          
          Dim tmpfile As LVZFileStruct
          Dim FramesX As Integer
          Dim FramesY As Integer
          Dim animTime As Long
          Dim filetitle As String
          
20        tmpfile = getFileFromNode(tvLVZfiles.SelectedItem)
          
30        If FileExists(tmpfile.path) Then
40            If tmpfile.Type = lvz_image Then
                  'special cases for animations
                  Dim previewimg As LVZImageDefinition
                  
50                previewimg.imagename = GetFileTitle(tmpfile.path)
                  
60                Call SetDefaultAnimationProperties(previewimg)
                      
70                Call ShowPreviewFile(tmpfile.path, previewimg.animationFramesX, previewimg.animationFramesY, previewimg.animationTime)
80            End If
90        End If
End Sub

Private Function getFileFromNode(Node As Node) As LVZFileStruct
          Dim ret As LVZFileStruct
10        Call lvz.setFileEmpty(ret)
          
20        If Node Is Nothing Then
              '...
30        ElseIf Node.parent Is Nothing Then
              '...
40        Else
50            ret = lvz.getFileData(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
60        End If
70        getFileFromNode = ret
End Function

Private Function getMapObjectFromNode(Node As Node) As LVZMapObject
          Dim ret As LVZMapObject
10        Call lvz.setMapObjectEmpty(ret)
20        If Node Is Nothing Then
              '...
30        ElseIf Node.parent Is Nothing Then
              '...
40        Else
50            ret = lvz.getMapObject(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
60        End If
70        getMapObjectFromNode = ret
End Function

Private Function getScreenObjectFromNode(Node As Node) As LVZScreenObject
          Dim ret As LVZScreenObject
10        Call lvz.setScreenObjectEmpty(ret)
20        If Node Is Nothing Then
              '...
30        ElseIf Node.parent Is Nothing Then
              '...
40        Else
50            ret = lvz.getScreenObject(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
60        End If
70        getScreenObjectFromNode = ret
End Function

Private Function getImageDefinitionFromMapObjNode(Node As Node) As LVZImageDefinition
          Dim ret As LVZImageDefinition
10        Call lvz.setImageDefinitionEmpty(ret)
20        If Node Is Nothing Then
              '...
30        ElseIf Node.parent Is Nothing Then
              '...
40        Else
              Dim tmpMapObj As LVZMapObject
50            tmpMapObj = getMapObjectFromNode(Node)
60            If tmpMapObj.imgidx <> -1 Then
70                ret = lvz.getImageDefinition(getIndexOfLVZFromNode(Node), tmpMapObj.imgidx)
80            End If
90        End If
100       getImageDefinitionFromMapObjNode = ret
End Function

Private Function getImageDefinitionFromScrObjNode(Node As Node) As LVZImageDefinition
          Dim ret As LVZImageDefinition
10        Call lvz.setImageDefinitionEmpty(ret)
20        If Node Is Nothing Then
              '...
30        ElseIf Node.parent Is Nothing Then
              '...
40        Else
              Dim tmpScrObj As LVZScreenObject
50            tmpScrObj = getScreenObjectFromNode(Node)
60            If tmpScrObj.imgidx <> -1 Then
70                ret = lvz.getImageDefinition(getIndexOfLVZFromNode(Node), tmpScrObj.imgidx)
80            End If
90        End If
100       getImageDefinitionFromScrObjNode = ret
End Function



Private Function getImageDefinitionFromNode(Node As Node) As LVZImageDefinition
          Dim ret As LVZImageDefinition
10        Call lvz.setImageDefinitionEmpty(ret)
20        If Node Is Nothing Then
              '...
30        ElseIf Node.parent Is Nothing Then
              '...
40        Else
50            ret = lvz.getImageDefinition(getIndexOfLVZFromNode(Node), ParseID(Node.Key))
60        End If
70        getImageDefinitionFromNode = ret
End Function

Private Function getIndexOfLVZFromNode(Node As Node) As Integer
10        If Node Is Nothing Then
20            getIndexOfLVZFromNode = -1
30        ElseIf Node.parent Is Nothing Then
40            getIndexOfLVZFromNode = -1
50        Else
60            getIndexOfLVZFromNode = lvz.getIndexOfLVZ(Node.parent.Text)
70        End If
End Function

Private Sub tvAvailableFiles_DblClick()

          Dim Node As Node
10        Set Node = tvImageDefinitions.SelectedItem
20        If Node Is Nothing Then Exit Sub
30        If Node.parent Is Nothing Then Exit Sub
          
          Dim imgNode As Node
40        Set imgNode = tvAvailableFiles.SelectedItem
50        If imgNode Is Nothing Then
60            Exit Sub
70        ElseIf imgNode.parent Is Nothing Then
80            Exit Sub
90        End If

          Dim idx As Integer
100       idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
110       If idx = -1 Then
              'parent LVZ not found
120           Exit Sub
130       End If
          
140       Call lvz.setImageDefinitionFile(idx, ParseID(Node.Key), imgNode.Text)
          
150       Call UpdateImageDefinitionListData
          
              
160       Node.Text = "Image" & ParseID(Node.Key) & IIf(imgNode.Text <> "", " (" & imgNode.Text & ")", " - No File")
          
End Sub

Private Sub tvLVZfiles_Expand(ByVal Node As MSComctlLib.Node)
10        Node.selected = True
End Sub

Private Sub tvLVZfiles_NodeClick(ByVal Node As MSComctlLib.Node)
10        Call tvLVZfiles_Click
End Sub

Private Sub tvLVZfiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim i As Integer, j As Integer
          Dim path As String
          Dim Node As Node
          Dim imported As Boolean
10        imported = False
          
20        For j = 1 To Data.files.count
30            path = Data.files(j)
40            If GetExtension(path) = "lvz" Then
50                If lvz.getIndexOfLVZ(GetFileTitle(path)) <> -1 Then
60                    MessageBox "Could not add lvz file, another file with the same name already exists.", vbOKOnly + vbExclamation
70                Else
80                    Call lvz.importLVZ(path, False)
90                    imported = True
                      
100                   Set Node = tvLVZfiles.Nodes.add(, , GetFileTitle(path), GetFileTitle(path), 3)
110               End If
120           Else
130               Set Node = tvLVZfiles.HitTest(X, Y)
140               If Node Is Nothing Then
                  
150                   If tvLVZfiles.Nodes.count = 0 Then
                          'add new lvz
                          Dim str As String
160                       i = lvz.getLVZCount
                          
170                       str = "Lvz" & i
                          
180                       While lvz.getIndexOfLVZ(str) <> -1
190                           i = i + 1
200                           str = "Lvz" & i
210                       Wend
                          
220                       If lvz.AddLVZ(str) Then
230                           Set Node = tvLVZfiles.Nodes.add(, , str, str, 3)
240                       End If
250                   End If
                      
260                   If lvz.getLVZCount > 0 Then
270                       Call AddFile(path, lvz.getLVZCount - 1)
280                   End If
                      
290               Else
300                   If Node.parent Is Nothing Then
310                       Call AddFile(path, lvz.getIndexOfLVZ(Node.Text))
320                   Else
330                       Call AddFile(path, lvz.getIndexOfLVZ(Node.parent.Text))
340                   End If
350               End If
                  
                  
              
360           End If
370       Next
          
380       If Node Is Nothing Then
          
390       Else
400           Node.selected = True
410           Node.Expanded = True
420       End If
430       If imported Then
440           Call RebuildFilesTree
450       End If
460       Call UpdateEnabledStatus
End Sub

Private Sub tvLVZfiles_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
          Dim Node As Node
10        Set Node = tvLVZfiles.HitTest(X, Y)
20        If Node Is Nothing Then Exit Sub
30        Node.selected = True
40        Node.Expanded = True
End Sub

Private Sub tvMapObjects_Click()
10        Call RefreshAvailableMapImages
20        Call UpdateMapObjectListData
30        Call UpdateEnabledStatus

40        Call showpreview(getImageDefinitionFromMapObjNode(tvMapObjects.SelectedItem))
          
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
10        If tvLVZfiles.SelectedItem Is Nothing Then
20            cmdRemoveItem.Enabled = False
30            cmdRemoveLVZ.Enabled = False
40            cmdExport.Enabled = False
50            cmdAddItem.Enabled = False
60            cmdAddLVZ.Enabled = True
70        ElseIf tvLVZfiles.SelectedItem.parent Is Nothing Then
80            cmdRemoveItem.Enabled = False
90            cmdRemoveLVZ.Enabled = True
100           cmdExport.Enabled = True
110           cmdAddItem.Enabled = True
120           cmdAddLVZ.Enabled = True
130       ElseIf Not tvLVZfiles.SelectedItem.parent Is Nothing Then
140           cmdRemoveItem.Enabled = True
150           cmdRemoveLVZ.Enabled = True
160           cmdExport.Enabled = True
170           cmdAddItem.Enabled = True
180           cmdAddLVZ.Enabled = True
190       End If

200       If tvMapObjects.SelectedItem Is Nothing Then
210           cmdRemoveMapObject.Enabled = False
220           cmdAddMapObject.Enabled = False
230           lstMapObjectProperties.Enabled = False
240       ElseIf tvMapObjects.SelectedItem.parent Is Nothing Then
250           cmdRemoveMapObject.Enabled = False
260           cmdAddMapObject.Enabled = True
270           lstMapObjectProperties.Enabled = False
280       ElseIf Not tvMapObjects.SelectedItem.parent Is Nothing Then
290           cmdRemoveMapObject.Enabled = True
300           cmdAddMapObject.Enabled = True
310           lstMapObjectProperties.Enabled = True
320           Call UpdateMapObjectListData
330       End If
          
340       If tvScreenObjects.SelectedItem Is Nothing Then
350           cmdRemoveScreenObject.Enabled = False
360           cmdAddScreenObject.Enabled = False
370           lstScreenObjectProperties.Enabled = False
380       ElseIf tvScreenObjects.SelectedItem.parent Is Nothing Then
390           cmdRemoveScreenObject.Enabled = False
400           cmdAddScreenObject.Enabled = True
410           lstScreenObjectProperties.Enabled = False
420       ElseIf Not tvScreenObjects.SelectedItem.parent Is Nothing Then
430           cmdRemoveScreenObject.Enabled = True
440           cmdAddScreenObject.Enabled = True
450           lstScreenObjectProperties.Enabled = True
460           Call UpdateScreenObjectListData
470       End If

480       If tvImageDefinitions.SelectedItem Is Nothing Then
490           cmdRemoveImgDef.Enabled = False
500           cmdAddImgDef.Enabled = False
510           lstImageDefinitionProperties.Enabled = False
520       ElseIf tvImageDefinitions.SelectedItem.parent Is Nothing Then
530           cmdRemoveImgDef.Enabled = False
540           cmdAddImgDef.Enabled = True
550           lstImageDefinitionProperties.Enabled = False
560       ElseIf Not tvImageDefinitions.SelectedItem.parent Is Nothing Then
570           cmdRemoveImgDef.Enabled = True
580           cmdAddImgDef.Enabled = True
590           lstImageDefinitionProperties.Enabled = True
600           Call UpdateImageDefinitionListData
610       End If
          
End Sub

Sub UpdateMapObjectListData()
          'lstMapObjectProperties.ListItems.Clear

          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
          'only show stuff if we have actually selected a map object
20        If Node Is Nothing Then
30            Exit Sub
40        ElseIf Node.parent Is Nothing Then
50            Exit Sub
60        End If
          
          Dim idx As Integer
70        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
80        If idx = -1 Then
              'parent LVZ not found
90            Exit Sub
100       End If
          
          
          Dim objidx As Long
110       objidx = ParseID(Node.Key)
          
120       If objidx = -1 Then Exit Sub
          
          Dim tmpMapObj As LVZMapObject
130       tmpMapObj = lvz.getMapObject(idx, objidx)
          
140       Call lstMapObjectProperties.setPropertyValue(LABEL_x, CStr(tmpMapObj.X))
150       Call lstMapObjectProperties.setPropertyValue(LABEL_y, CStr(tmpMapObj.Y))
160       If val(tmpMapObj.imgidx) = -1 Then
170           Call lstMapObjectProperties.setPropertyValue(LABEL_image, "")
180       Else
190           Call lstMapObjectProperties.setPropertyValue(LABEL_image, "Image" & tmpMapObj.imgidx & " (" & lvz.getImageDefinition(idx, tmpMapObj.imgidx).imagename & ")")
200       End If
210       Call lstMapObjectProperties.setPropertyValue(LABEL_layer, CStr(tmpMapObj.layer))
220       Call lstMapObjectProperties.setPropertyValue(LABEL_mode, CStr(tmpMapObj.mode))
          
230       Call lstMapObjectProperties.setPropertyValue(LABEL_displayTime, CStr(tmpMapObj.displayTime))
240       Call lstMapObjectProperties.setPropertyToolTipText(LABEL_displayTime, "Time that the image should be displayed in 1/10th of seconds. Has no effect if Display Mode is set to 'Show Always'")
          
250       Call lstMapObjectProperties.setPropertyValue(LABEL_objectID, CStr(tmpMapObj.objectID))
260       Call lstMapObjectProperties.setPropertyToolTipText(LABEL_objectID, "*objon #")
End Sub

Sub UpdateScreenObjectListData()
          'lstMapObjectProperties.ListItems.Clear

          Dim Node As Node
10        Set Node = tvScreenObjects.SelectedItem
          'only show stuff if we have actually selected a map object
20        If Node Is Nothing Then
30            Exit Sub
40        ElseIf Node.parent Is Nothing Then
50            Exit Sub
60        End If
          
          Dim idx As Integer
70        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
80        If idx = -1 Then
              'parent LVZ not found
90            Exit Sub
100       End If
          
          
          Dim objidx As Long
110       objidx = ParseID(Node.Key)
          
120       If objidx = -1 Then Exit Sub
          
          Dim tmpscreenObj As LVZScreenObject
130       tmpscreenObj = lvz.getScreenObject(idx, objidx)
          
140       Call lstScreenObjectProperties.setPropertyValue(LABEL_x, CStr(tmpscreenObj.X))
150       Call lstScreenObjectProperties.setPropertyValue(LABEL_y, CStr(tmpscreenObj.Y))
160       Call lstScreenObjectProperties.setPropertyValue(LABEL_typeX, CStr(tmpscreenObj.typeX))
170       Call lstScreenObjectProperties.setPropertyValue(LABEL_typeY, CStr(tmpscreenObj.typeY))
180       If val(tmpscreenObj.imgidx) = -1 Then
190           Call lstScreenObjectProperties.setPropertyValue(LABEL_image, "")
200       Else
210           Call lstScreenObjectProperties.setPropertyValue(LABEL_image, "Image" & tmpscreenObj.imgidx & " (" & lvz.getImageDefinition(idx, tmpscreenObj.imgidx).imagename & ")")
220       End If
230       Call lstScreenObjectProperties.setPropertyValue(LABEL_layer, CStr(tmpscreenObj.layer))
240       Call lstScreenObjectProperties.setPropertyValue(LABEL_mode, CStr(tmpscreenObj.mode))
250       Call lstScreenObjectProperties.setPropertyValue(LABEL_displayTime, CStr(tmpscreenObj.displayTime))
260       Call lstScreenObjectProperties.setPropertyValue(LABEL_objectID, CStr(tmpscreenObj.objectID))
End Sub

Sub UpdateImageDefinitionListData()
          'lstMapObjectProperties.ListItems.Clear

          Dim Node As Node
10        Set Node = tvImageDefinitions.SelectedItem
          'only show stuff if we have actually selected a map object
20        If Node Is Nothing Then
30            Exit Sub
40        ElseIf Node.parent Is Nothing Then
50            Exit Sub
60        End If
          
          Dim idx As Integer
70        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
80        If idx = -1 Then
              'parent LVZ not found
90            Exit Sub
100       End If
          
          
          Dim objidx As Integer
110       objidx = ParseID(Node.Key)
          
120       If objidx = -1 Then Exit Sub
          
          Dim tmpImgDef As LVZImageDefinition
130       tmpImgDef = lvz.getImageDefinition(idx, objidx)
          
140       Call lstImageDefinitionProperties.setPropertyValue(LABEL_imagename, CStr(tmpImgDef.imagename))
150       Call lstImageDefinitionProperties.setPropertyValue(LABEL_animationFramesX, CStr(tmpImgDef.animationFramesX))
160       Call lstImageDefinitionProperties.setPropertyValue(LABEL_animationFramesY, CStr(tmpImgDef.animationFramesY))
170       Call lstImageDefinitionProperties.setPropertyValue(LABEL_animationTime, CStr(tmpImgDef.animationTime))
End Sub

Private Sub tvMapObjects_DblClick()
          Dim Node As Node
10        Set Node = tvMapObjects.SelectedItem
          'only show stuff if we have actually selected a map object
20        If Node Is Nothing Then
30            Exit Sub
40        ElseIf Node.parent Is Nothing Then
50            Exit Sub
60        End If
          
          Dim idx As Integer
70        idx = lvz.getIndexOfLVZ(Node.parent.Text)
          
80        If idx = -1 Then
              'parent LVZ not found
90            Exit Sub
100       End If
          
          
          Dim objidx As Long
110       objidx = ParseID(Node.Key)
          
120       If objidx = -1 Then Exit Sub
          
          Dim tmpMapObj As LVZMapObject
130       tmpMapObj = lvz.getMapObject(idx, objidx)
          
          
140       Call parent.SetFocusAt(tmpMapObj.X \ 16, tmpMapObj.Y \ 16, parent.picPreview.width \ 2, parent.picPreview.height \ 2, True)

End Sub

Private Sub tvMapObjects_NodeClick(ByVal Node As MSComctlLib.Node)
10        Call tvMapObjects_Click
End Sub

Private Sub tvScreenObjects_Click()
10        Call RefreshAvailableScreenImages
20        Call UpdateScreenObjectListData
30        Call UpdateEnabledStatus
          
40        Call showpreview(getImageDefinitionFromScrObjNode(tvScreenObjects.SelectedItem))
End Sub

Private Sub ShowPreviewFile(path As String, Optional FramesX As Integer = 1, Optional FramesY As Integer = 1, Optional animTime As Integer = 100)
          Const BORDER_LEFT = 8
          Const BORDER_TOP = 16
          Const BORDER_RIGHT = 8
          Const BORDER_BOTTOM = 8
          
10        If Not FileExists(path) Then Exit Sub
          
20        TimerAnim.Enabled = False
          
30        Call ClearPreview
          
40        previewFramesX = FramesX
50        previewFramesY = FramesY
          
60        picPreview.AutoSize = True
70        Call LoadPic(picPreview, path)
80        picPreview.Refresh
90        picPreview.AutoSize = False

          Dim picidx As Integer
100       picidx = tblvz.SelectedItem.Index - 1
          
          Dim trueWidth As Integer
          Dim trueHeight As Integer
110       trueWidth = picPreview.width \ previewFramesX
120       trueHeight = picPreview.height \ previewFramesY
          
          Dim maxWidth As Integer
          Dim maxHeight As Integer
130       maxWidth = picPreviewAnim(picidx).Container.width - BORDER_LEFT - BORDER_RIGHT
140       maxHeight = picPreviewAnim(picidx).Container.height - BORDER_TOP - BORDER_BOTTOM
          
          Dim newWidth As Integer
          Dim newHeight As Integer
              
150       If trueWidth > maxWidth Or _
             trueHeight > maxHeight Then
              Dim ratio As Double
160           ratio = doubleMinimum(maxWidth / trueWidth, maxHeight / trueHeight)
170           newWidth = Int(trueWidth * ratio)
180           newHeight = Int(trueHeight * ratio)
190       Else
              'center it
200           newWidth = trueWidth
210           newHeight = trueHeight
220       End If
              
230       picPreviewAnim(picidx).width = newWidth * Screen.TwipsPerPixelX
240       picPreviewAnim(picidx).height = newHeight * Screen.TwipsPerPixelY
250       picPreviewAnim(picidx).Left = (BORDER_LEFT + (maxWidth - newWidth) \ 2) * Screen.TwipsPerPixelX
260       picPreviewAnim(picidx).Top = (BORDER_TOP + (maxHeight - newHeight) \ 2) * Screen.TwipsPerPixelY
270       picPreviewAnim(picidx).visible = True

280       TimerAnim.Interval = (animTime / (FramesX * FramesY)) * 10
          
290       previewFileName = GetFileTitle(path)
          
300       Call RefreshPreview(True, True)
          
310       TimerAnim.Enabled = True
End Sub

Private Sub showpreview(imgdef As LVZImageDefinition)

          Dim path As String
10        path = lvz.SearchFile(imgdef.imagename)
          
20        If Not FileExists(path) Then Exit Sub
          
30        Call ShowPreviewFile(path, imgdef.animationFramesX, imgdef.animationFramesY, imgdef.animationTime)
          
End Sub

Private Sub ClearPreview()
          Dim i As Integer
10        For i = 0 To 3
20            picPreviewAnim(i).Cls
30            picPreviewAnim(i).visible = False
40        Next
End Sub

Private Sub RefreshPreview(resetFrame As Boolean, changeFrame As Boolean)

          
          Static frameID As Long
          
10        If resetFrame Then
20            frameID = 0
30        End If
          
          Dim frameX As Integer
          Dim frameY As Integer
40        frameX = (frameID Mod previewFramesX)
50        frameY = ((frameID \ previewFramesX) Mod previewFramesY)

          Dim picidx As Integer
60        picidx = tblvz.SelectedItem.Index - 1
          
          Dim trueWidth As Integer
          Dim trueHeight As Integer
70        trueWidth = picPreview.width \ previewFramesX
80        trueHeight = picPreview.height \ previewFramesY
              
90        fPreview(picidx).Caption = "Preview: " & previewFileName & " - " & trueWidth & Chr(215) & trueHeight & " pixels - Frame " & (previewFramesX * frameY + frameX) + 1 & "/" & previewFramesX * previewFramesY
                          
100       If trueWidth <> picPreviewAnim(picidx).ScaleWidth Or _
             trueHeight > picPreviewAnim(picidx).ScaleHeight Then
              'stretchblt
110           SetStretchBltMode picPreviewAnim(picidx).hDC, HALFTONE
120           Call StretchBlt(picPreviewAnim(picidx).hDC, 0, 0, picPreviewAnim(picidx).ScaleWidth, picPreviewAnim(picidx).ScaleHeight, picPreview.hDC, trueWidth * frameX, trueHeight * frameY, trueWidth, trueHeight, vbSrcCopy)
130       Else
              'bitblt
140           Call BitBlt(picPreviewAnim(picidx).hDC, 0, 0, trueWidth, trueHeight, picPreview.hDC, trueWidth * frameX, trueHeight * frameY, vbSrcCopy)
150       End If

160       picPreviewAnim(picidx).Refresh
          
170       If changeFrame Then frameID = (frameID + 1) Mod 1073676289 '32767 * 32767, max # of frames possible

End Sub

Private Sub tvScreenObjects_NodeClick(ByVal Node As MSComctlLib.Node)
10        Call tvScreenObjects_Click
End Sub
