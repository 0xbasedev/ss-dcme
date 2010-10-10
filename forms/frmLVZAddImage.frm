VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLVZAddImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Image To Library"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   509
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "File Settings"
      Height          =   1455
      Left            =   3600
      TabIndex        =   25
      Top             =   2160
      Width           =   3975
      Begin VB.ComboBox cmbFileLVZ 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbDefLVZ 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Place Image File In:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Place Image Definition In:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame frmPreview 
      Caption         =   "Preview - 1:1"
      Height          =   3255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   7455
      Begin VB.CommandButton cmdzoomin 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdzoomout 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   0
         Width           =   375
      End
      Begin DCME.cPicViewer picPreview 
         Height          =   2655
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   7215
         _ExtentX        =   6588
         _ExtentY        =   6376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AnimationTime   =   1000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Source"
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   7455
      Begin VB.PictureBox picSource 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   1
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   6855
         TabIndex        =   23
         Top             =   1560
         Width           =   6855
         Begin VB.ComboBox cmbSourceImage 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cmbSourceLVZ 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Image File:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "LVZ File:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox picSource 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   2760
         ScaleHeight     =   495
         ScaleWidth      =   3975
         TabIndex        =   22
         Top             =   1440
         Width           =   3975
      End
      Begin VB.PictureBox picSource 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   0
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   7215
         TabIndex        =   14
         Top             =   1200
         Width           =   7215
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   0
            Width           =   5295
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   6000
            TabIndex        =   15
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "File:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Import Image From File"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Import an image file (bmp, png, jpg, gif) to create a new image definition."
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Use Existing Image In Library"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Use the same file as another image in your lvzs. Use this to create multiple different animations of the same image, for example."
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Create New Image"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame frmAnimation 
      Caption         =   "Animation Settings"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox txtAnimTime 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "100"
         ToolTipText     =   "Time in hundreths of seconds of the whole animation loop"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtFramesY 
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Number of frames vertically"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtFramesX 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Text            =   "1"
         ToolTipText     =   "Number of frames horizontally"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkAnimation 
         Caption         =   "Animation"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Animation Loop Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Frames:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   7080
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   480
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmLVZAddImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lvz As LVZData
Dim parent As frmMain

Dim imageselected As Boolean


'Where is the image taken from
Private Enum enumImageSource
    FromFile = 0
    ExistingFile = 1
    NewImage = 2
End Enum

Dim Source As enumImageSource


Private Sub chkAnimation_Click()
    Call CheckAnimation
End Sub

Private Sub CheckAnimation()
    Dim value As Boolean
    
    value = (chkAnimation.value = vbChecked)
    
    
    txtAnimTime.Enabled = value
    txtAnimTime.Text = picPreview.animationTime \ 10
    
    txtFramesX.Enabled = value
    txtFramesX.Text = picPreview.AnimFramesX
    
    txtFramesY.Enabled = value
    txtFramesY.Text = picPreview.AnimFramesY
    
    picPreview.Animation = value
    
End Sub

Private Sub cmbSourceImage_Click()
    Dim itemdata As Long
    itemdata = cmbSourceImage.itemdata(cmbSourceImage.ListIndex)
    
    If itemdata = -1 Then
        Call picPreview.Clear
        cmdAdd.Enabled = False
        
    Else
        Dim lvzidx As Integer, imgidx As Long
        lvzidx = itemdata \ 65536
        imgidx = itemdata Mod 65536
        
        Call LoadPreviewPic(lvz.getFileData(lvzidx, imgidx).path)
    End If
End Sub

Private Sub AddImagesOfLVZ(lvzidx As Integer)
    Dim count As Integer, i As Long
    
    count = lvz.getFileCount(lvzidx)
    For i = 0 To count - 1
        Dim File As LVZFileStruct
        
        File = lvz.getFileData(lvzidx, i)
        
        If File.Type = lvz_image Then
            'Add it
            Call cmbSourceImage.addItem(GetFileTitle(File.path))
            
            
            'lvzidx = data \ 65536
            'fileidx = data Mod 65536
            cmbSourceImage.itemdata(cmbSourceImage.NewIndex) = i + lvzidx * 65536

        End If
    Next
End Sub


Private Sub cmbSourceLVZ_Click()
    Dim i As Integer
    Dim count As Integer
    Dim itemstr As String

    If cmbSourceLVZ.ListIndex = -1 Then
        cmbSourceLVZ.ListIndex = 0
    End If
    
    cmbSourceImage.Clear
    
    itemstr = cmbSourceLVZ.list(cmbSourceLVZ.ListIndex)
    
    If itemstr = "* SHOW ALL *" Then
        'Show all, duh
        count = lvz.getLVZCount
        For i = 0 To count - 1
            Call AddImagesOfLVZ(i)
        Next
    Else
        Dim idx As Integer
        idx = lvz.getIndexOfLVZ(itemstr)
        If idx = -1 Then
            MessageBox "Error! LVZ '" & itemstr & "' not found."
        Else
            'Search for all images in that LVZ
            Call AddImagesOfLVZ(idx)
        End If
    End If

    If cmbSourceImage.ListCount = 0 Then
        cmbSourceImage.addItem "* NONE *"
        cmbSourceImage.itemdata(0) = -1
        
        cmbSourceImage.ListIndex = 0
        cmbSourceImage.Enabled = False
        imageselected = False
    Else
        cmbSourceImage.Enabled = True
        cmbSourceImage.ListIndex = 0
        
        Call cmbSourceImage_Click
    End If
    
End Sub

Private Function PromptNewLvz(prompt As String) As String
    Dim name As String
    Do
        name = InputBox(prompt, "New LVZ file", "Lvz" & lvz.getLVZCount)
        
        If name = "" Then MessageBox "You must enter a name for this LVZ file!"
        
    Loop While name = ""
    
    If GetExtension(name) <> "lvz" Then name = name & ".lvz"
    
    PromptNewLvz = Trim$(replace$(name, "\", "_"))
End Function

Private Function CreateNewLvz(name As String) As Integer
    'Create new lvz
    If lvz.AddLVZ(name) Then
        CreateNewLvz = lvz.getIndexOfLVZ(name)
    Else
        'Name already exists
        'Error
        CreateNewLvz = -1
    End If
End Function


Private Sub cmdAdd_Click()

    If Not imageselected Then
        MessageBox "No image selected!", vbExclamation
        Exit Sub
    End If
    
    Dim i As Integer
    Dim filelvzidx As Integer, deflvzidx As Integer, imgidx As Integer, fileIdx As Integer
    
    Dim filelvz As String, deflvz As String, newlvzname As String
    
    filelvzidx = -1
    deflvzidx = -1
    imgidx = -1
    fileIdx = -1
    
    If Source = ExistingFile Then
        Dim imgsrc As Long
        imgsrc = cmbSourceImage.itemdata(cmbSourceImage.ListIndex)
        
        filelvzidx = imgsrc \ 65536
        fileIdx = imgsrc Mod 65536
        
        
    ElseIf Source = FromFile Then
        filelvz = cmbFileLVZ.list(cmbFileLVZ.ListIndex)
        
        If filelvz = "* NEW LVZ *" Then
            'Ask for new lvz's name
            newlvzname = PromptNewLvz("Image file LVZ")
            
            filelvzidx = CreateNewLvz(newlvzname)
            
            If filelvzidx = -1 Then
                MessageBox "LVZ file '" & newlvzname & "' already exists!", vbExclamation
                Exit Sub
            End If
        Else
            filelvzidx = lvz.getIndexOfLVZ(filelvz)
        End If
        
        If filelvzidx = -1 Then
            'Error
            MessageBox "LVZ file '" & filelvz & "' was not found", vbExclamation
            Exit Sub
        ElseIf FileExists(txtPath.Text) Then
            
            fileIdx = lvz.AddFileToLVZ(filelvzidx, txtPath.Text)
            
            If fileIdx = -1 Then
                MessageBox "File '" & GetFileTitle(txtPath.Text) & "' already exists in LVZ!", vbExclamation
                Exit Sub
            End If
            
        Else
            MessageBox "File '" & txtPath.Text & "' does not exist!", vbExclamation
            Exit Sub
        End If
    End If
    
    'By now, we should have the index of the lvz, and of the image file
    
    'Find where to put the new image definition
    deflvz = cmbDefLVZ.list(cmbDefLVZ.ListIndex)
    
    If deflvz = "* NEW LVZ *" Then
        If filelvz <> "* NEW LVZ *" Then
            'Create a new LVZ for it
            newlvzname = PromptNewLvz("Image definition LVZ")
           
            deflvzidx = CreateNewLvz(newlvzname)
            
            If deflvzidx = -1 Then
                MessageBox "LVZ file '" & newlvzname & "' already exists!", vbExclamation
                Exit Sub
            End If
        Else
            'Use the same new lvz than the file
            deflvzidx = filelvzidx
        End If
    Else
        deflvzidx = lvz.getIndexOfLVZ(deflvz)
        
        If deflvzidx = -1 Then
            'Error
            MessageBox "LVZ file '" & deflvz & "' was not found", vbExclamation
            Exit Sub
        End If
    End If
    
    'Now let's make ourselves an image definition
    Dim imgdef As LVZImageDefinition
    
    If Source = ExistingFile Then
        imgdef.imagename = cmbSourceImage.list(cmbSourceImage.ListIndex)
    ElseIf Source = FromFile Then
        imgdef.imagename = GetFileTitle(txtPath.Text)
    End If
    
    If chkAnimation.value = vbChecked Then
        imgdef.animationTime = val(txtAnimTime.Text)
        imgdef.animationFramesX = val(txtFramesX.Text)
        imgdef.animationFramesY = val(txtFramesY.Text)
    Else
        imgdef.animationFramesX = 1
        imgdef.animationFramesY = 1
    End If
    
    imgidx = lvz.AddImageDefinitionToLVZ(deflvzidx, imgdef)
    
    If imgidx = -1 Then
        MessageBox "Maximum number of image definitions reached (256)!", vbExclamation
        Exit Sub
    End If
    
    Me.MousePointer = MousePointerConstants.vbHourglass
    Call lvz.buildAllLVZImages
    
    Call InitForm
    Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub cmdzoomin_Click()
    Call picPreview.ZoomIn
End Sub

Private Sub cmdzoomout_Click()
    Call picPreview.ZoomOut
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon
    
    'Place pictureboxes correctly
    Dim fx As Integer, fy As Integer
    fx = picSource(0).Left
    fy = picSource(0).Top
    
    Dim fpic As PictureBox
    For Each fpic In picSource
        fpic.Left = fx
        fpic.Top = fy
    Next
    
    Call InitForm
    
    
End Sub

Private Sub InitForm()
    txtPath.Text = ""
    cmbFileLVZ.Clear
    cmbDefLVZ.Clear
    cmbSourceLVZ.Clear
    cmbSourceImage.Clear
    
    cmbFileLVZ.addItem "* NEW LVZ *"
    cmbDefLVZ.addItem "* NEW LVZ *"
    cmbSourceLVZ.addItem "* SHOW ALL *"
    
    'Populate LVZ files
    Dim lvzs() As LVZstruct
    lvzs = lvz.getLVZData
    
    Dim lvzcount As Integer, imgCount As Integer
    
    lvzcount = lvz.getLVZCount
    
    Dim i As Integer
    For i = 0 To lvzcount - 1
        cmbFileLVZ.addItem lvzs(i).name
        cmbDefLVZ.addItem lvzs(i).name
        cmbSourceLVZ.addItem lvzs(i).name
    Next
    
    'By default, select the last added item
    cmbFileLVZ.ListIndex = cmbFileLVZ.ListCount - 1
    cmbDefLVZ.ListIndex = cmbDefLVZ.ListCount - 1
    'cmbSourceLVZ.ListIndex = cmbSourceLVZ.ListCount - 1
    
    cmdAdd.Enabled = False
    
    txtAnimTime.Text = 100
    picPreview.animationTime = 1000
    
    Call CheckAnimation
    imageselected = False
    optSource(0).value = True
    
    picPreview.Clear
End Sub

Public Sub setParent(Main As frmMain)
    Set parent = Main
    Set lvz = Main.lvz
End Sub

Private Sub LoadPreviewPic(path As String)
    If path = "" Or Not FileExists(path) Then
        imageselected = False
        cmdAdd.Enabled = False
        picPreview.Clear
        txtPath.Text = "NO FILE LOADED"
    Else
        'txtPath.Text = path
        Call picPreview.LoadPicture(path)
    
        imageselected = True
        cmdAdd.Enabled = True
    End If
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo errh
    'opens a common dialog
    cd.DialogTitle = "Select an image to import"
    cd.flags = cdlOFNHideReadOnly

    cd.InitDir = GetLastDialogPath("AddLvzImage")
    
    cd.Filter = "Supported image files (*.lvl, *.bmp, *.png, *.gif, *.jpg)|*.lvl; *.bmp; *.bm2; *.gif; *.jpg; *.jpeg; *.png|Tilesets (*.lvl, *.bmp)|*.lvl; *.bmp|All files (*.*)|*.*"
    cd.ShowOpen
    
    txtPath.Text = cd.filename
    Call LoadPreviewPic(cd.filename)
    
    Call SetLastDialogPath("AddLvzImage", GetPathTo(cd.filename))
    
    Exit Sub
errh:
    'if something goes wrong, use the default tileset
    If Err = cdlCancel Then
        Exit Sub
    Else
        MessageBox Err & " " & Err.description, vbCritical
    End If

    On Error GoTo 0
    Exit Sub
End Sub






Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Set lvz = Nothing
    Set parent = Nothing
End Sub

Private Sub optSource_Click(Index As Integer)
    Dim pframe As PictureBox
    
    For Each pframe In picSource
        pframe.visible = (pframe.Index = Index)
    Next
    
    Source = Index
    
    If Source = ExistingFile Then
        'Existing image file
        cmbFileLVZ.Enabled = False
    Else
        cmbFileLVZ.Enabled = True
    End If
    
    Call picPreview.Clear
    cmdAdd.Enabled = False
    
    If Source = ExistingFile Then
        Call cmbSourceLVZ_Click
        
    ElseIf Source = FromFile Then
        If FileExists(txtPath.Text) Then
            Call LoadPreviewPic(txtPath.Text)
        End If
    End If
End Sub



Private Sub picPreview_Zoom()
    frmPreview.caption = "Preview - " & picPreview.zoomstr
End Sub



Private Sub txtAnimTime_Change()
    Call removeDisallowedCharacters(txtAnimTime, 1, 32767, False)
    
    picPreview.animationTime = val(txtAnimTime.Text) * 10
End Sub

Private Sub txtFramesX_Change()
    Call removeDisallowedCharacters(txtFramesX, 1, 32767, False)
    
    picPreview.AnimFramesX = val(txtFramesX)
End Sub

Private Sub txtFramesY_Change()
    Call removeDisallowedCharacters(txtFramesY, 1, 32767, False)
    
    picPreview.AnimFramesY = val(txtFramesY)
End Sub
