VERSION 5.00
Begin VB.Form frmEditRegion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Region"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picmap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   2040
      MousePointer    =   2  'Cross
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
      Begin VB.Shape cursor 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         Height          =   105
         Left            =   720
         Shape           =   3  'Circle
         Top             =   480
         Width           =   105
      End
      Begin VB.Line cursorLine2 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         X1              =   32
         X2              =   32
         Y1              =   56
         Y2              =   48
      End
      Begin VB.Line cursorLine1 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         X1              =   32
         X2              =   44
         Y1              =   56
         Y2              =   56
      End
      Begin VB.Line centerXLine 
         BorderColor     =   &H00FF0000&
         X1              =   0
         X2              =   129
         Y1              =   64
         Y2              =   64
      End
      Begin VB.Line centerYLine 
         BorderColor     =   &H00FF0000&
         X1              =   64
         X2              =   64
         Y1              =   32
         Y2              =   161
      End
   End
   Begin VB.CheckBox chkNoReceiveWeapons 
      Caption         =   "No Receive Weapons"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CheckBox chkNoReceiveAnti 
      Caption         =   "No Receive Antiwarp"
      Height          =   255
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   18
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox txtPython 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmEditRegion.frx":0000
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdEditPython 
      Caption         =   "Edit..."
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Jump To"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtArena 
      Height          =   285
      Left            =   2520
      MaxLength       =   16
      TabIndex        =   13
      ToolTipText     =   "Name of the arena to warp to. Leave blank for current arena."
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtAutowarpY 
      Height          =   285
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "512"
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtAutowarpX 
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "512"
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CheckBox chkAutowarp 
      Caption         =   "Autowarp"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox chkNoWeapons 
      Caption         =   "No Weapons"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox chkNoFlagDrops 
      Caption         =   "No Flag Drops"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CheckBox chkNoAntiWarp 
      Caption         =   "No Anti-Warp"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox chkIsBase 
      Caption         =   "Base"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Python Code"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblArena 
      Caption         =   "Arena:"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      ToolTipText     =   "Name of the arena to warp to. Leave blank for current arena"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblY 
      Caption         =   "Y:"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   1230
      Width           =   255
   End
   Begin VB.Label lblX 
      Caption         =   "X:"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   1230
      Width           =   255
   End
End
Attribute VB_Name = "frmEditRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const WIDTH1 = 4155    'width of form with autowarp Off
'Const WIDTH2 = 4155    'width of form with autowarp On
'Const WIDTH3 = 4155    'width of form when viewing the minimap to select a coordinate

'Const HEIGHT1 = 2940
'Const HEIGHT2 = 3735
'Const HEIGHT3 = 3735
Dim parent As frmMain
Dim Region As Region
Dim rIdx As Integer
Dim ignoreupdate As Boolean

Sub setParent(Main As frmMain, regionindex As Integer)
    Set parent = Main
    rIdx = regionindex

    Call LoadRegion
    Call UpdatePreview
End Sub

Private Sub chkAutowarp_Click()
    CheckAutowarpType
End Sub



Private Sub cmdEditPython_Click()
    frmPython.show vbModal, frmGeneral
    Region.pythonCode = frmGeneral.GetActiveRegionPythonCode
    If Region.pythonCode <> "" Then
        txtPython = "..."
    Else
        txtPython = "<None>"
    End If
End Sub

Private Sub cmdGoto_Click()
    Dim tileX As Integer, tileY As Integer

    tileX = val(txtAutowarpX.Text)
    tileY = val(txtAutowarpY.Text)

    If tileX = 0 Or tileY = 0 Then
        MessageBox "Invalid jump coordinates. 0 is used for arena's default warp coordinate.", vbOKOnly + vbExclamation, "Invalid coordinate"
    ElseIf tileX = -1 Or tileY = -1 Then
        MessageBox "Invalid jump coordinates. -1 is used for player's current position.", vbOKOnly + vbExclamation, "Invalid coordinate"
    Else
        Call parent.SetFocusAt(tileX - 1, tileY - 1, parent.picPreview.width \ 2, parent.picPreview.height \ 2, True)
    End If
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    Dim frmtoPt As POINTAPI

    ClientToScreen frmGeneral.tlbToolOptions.hWnd, frmtoPt
    
'    Me.width = WIDTH3
'    Me.height = HEIGHT3

    Me.Left = frmtoPt.X * Screen.TwipsPerPixelX + frmGeneral.llRegionList.Left + GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
    Me.Top = frmtoPt.Y * Screen.TwipsPerPixelY + frmGeneral.llRegionList.height + (GetSystemMetrics(SM_CYMENU) + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFRAME)) * Screen.TwipsPerPixelY
    
    Call ShowQuickMap
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If picmap.visible Then
'        Call HideQuickMap
'    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveRegion
    Unload Me
End Sub

Private Sub LoadRegion()
    Set Region = parent.Regions.getRegion(rIdx)

    If Region.isBase Then
        chkIsBase.value = vbChecked
    Else
        chkIsBase.value = vbUnchecked
    End If

    If Region.isNoAntiwarp Then
        chkNoAntiWarp.value = vbChecked
    Else
        chkNoAntiWarp.value = vbUnchecked
    End If

    If Region.isNoWeapon Then
        chkNoWeapons.value = vbChecked
    Else
        chkNoWeapons.value = vbUnchecked
    End If

    If Region.isNoFlagDrop Then
        chkNoFlagDrops.value = vbChecked
    Else
        chkNoFlagDrops.value = vbUnchecked
    End If

    If Region.isAutoWarp Then
        chkAutowarp.value = vbChecked
    Else
        chkAutowarp.value = vbUnchecked
    End If


  If Region.isNoReceiveAntiwarp Then
      chkNoReceiveAnti.value = vbChecked
  Else
      chkNoReceiveAnti.value = vbUnchecked
  End If
  
  If Region.isNoReceiveWeapon Then
      chkNoReceiveWeapons.value = vbChecked
  Else
      chkNoReceiveWeapons.value = vbUnchecked
  End If
  
  
    If Region.pythonCode <> "" Then
        txtPython = "..."
    Else
        txtPython = "<None>"
    End If
    txtName.Text = Region.name
    
    txtAutowarpX.Text = Region.autowarpX
    txtAutowarpY.Text = Region.autowarpY

    txtArena.Text = Region.autowarpArena

    Call CheckAutowarpType

End Sub

Private Sub SaveRegion()
    If txtName.Text <> "" Then
        Region.name = txtName.Text
    End If
    Region.autowarpArena = txtArena.Text
    Region.autowarpX = val(txtAutowarpX.Text)
    Region.autowarpY = val(txtAutowarpY.Text)
    Region.isAutoWarp = (chkAutowarp.value = vbChecked)
    Region.isBase = (chkIsBase.value = vbChecked)
    Region.isNoAntiwarp = (chkNoAntiWarp.value = vbChecked)
    Region.isNoFlagDrop = (chkNoFlagDrops.value = vbChecked)
    Region.isNoWeapon = (chkNoWeapons.value = vbChecked)

  Region.isNoReceiveAntiwarp = (chkNoReceiveAnti.value = vbChecked)
  Region.isNoReceiveWeapon = (chkNoReceiveWeapons.value = vbChecked)



    'the python code was already saved in region
    Region.pythonCode = frmGeneral.GetActiveRegionPythonCode
    
'    Call parent.Regions.setRegion(Region, rIdx)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
    Set Region = Nothing
End Sub

Private Sub txtAutowarpX_Change()
    Dim newX As Integer
    Call removeDisallowedCharacters(txtAutowarpX, -1, 1024)

    newX = val(txtAutowarpX.Text)
    If newX = 0 Or newX = -1 Then
        newX = 512
    End If
    Call MoveCursorToX(newX)
End Sub

Private Sub txtAutowarpX_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAutowarpX.Text) Then
        txtAutowarpX.Text = 0
    Else
        txtAutowarpX.Text = val(txtAutowarpX.Text)
    End If
End Sub

Private Sub txtAutowarpY_Change()
    Dim newY As Integer
    Call removeDisallowedCharacters(txtAutowarpY, -1, 1024)

    newY = val(txtAutowarpY.Text)
    If newY = 0 Or newY = -1 Then
        newY = 512
    End If
    Call MoveCursorToY(newY)
End Sub



Private Sub ShowQuickMap()
    cmdGoto.visible = True
    picmap.visible = True
'    Me.width = WIDTH3
'    Me.height = HEIGHT3

End Sub

Private Sub HideQuickMap()
    cmdGoto.visible = False
    picmap.visible = False
'    Me.width = WIDTH2
'    Me.height = HEIGHT2
End Sub



Private Sub CheckAutowarpType()
    If chkAutowarp.value = vbChecked Then
        txtAutowarpX.Enabled = True
        txtAutowarpY.Enabled = True
        lblX.Enabled = True
        lblY.Enabled = True
        txtArena.Enabled = True
        lblArena.Enabled = True

'80            Me.width = WIDTH2
'90            Me.height = HEIGHT2
        Call ShowQuickMap


    Else
        txtAutowarpX.Enabled = False
        txtAutowarpY.Enabled = False
        lblX.Enabled = False
        lblY.Enabled = False
        txtArena.Enabled = False
        lblArena.Enabled = False

        Call HideQuickMap

'190           Me.width = WIDTH1
'200           Me.height = HEIGHT1
    End If
End Sub


Private Sub picmap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X <= 0 Then X = 0
    If X >= picmap.ScaleWidth - 1 Then X = picmap.ScaleWidth - 1
    If Y <= 0 Then Y = 0
    If Y >= picmap.ScaleHeight - 1 Then Y = picmap.ScaleHeight - 1

    Call PlacePointer(X, Y)
    Call UpdateValues
End Sub

Private Sub picmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        Call picmap_MouseDown(Button, Shift, X, Y)
    End If

End Sub

Private Sub PlacePointer(X As Single, Y As Single)
    Call PlacePointerX(X)
    Call PlacePointerY(Y)
End Sub

Private Sub PlacePointerX(X As Single)
    cursor.Left = X - 3

    cursorLine1.x1 = X
    cursorLine1.x2 = X

    cursorLine2.x1 = X - 5
    cursorLine2.x2 = X + 5
End Sub

Private Sub PlacePointerY(Y As Single)
    cursor.Top = Y - 3

    cursorLine1.y1 = Y - 5
    cursorLine1.y2 = Y + 5

    cursorLine2.y1 = Y
    cursorLine2.y2 = Y
End Sub


Private Sub UpdateValues()
    ignoreupdate = True
    txtAutowarpX.Text = Int(((cursor.Left + 3) / (picmap.width - 1)) * 1023) + 1
    txtAutowarpY.Text = Int(((cursor.Top + 3) / (picmap.height - 1)) * 1023) + 1
    ignoreupdate = False
End Sub

Private Sub MoveCursorToX(Xval As Integer)
    If ignoreupdate Then Exit Sub

    Dim X As Single
    X = Int(((Xval - 1) / 1024) * (picmap.width - 1))

    Call PlacePointerX(X)
End Sub

Private Sub MoveCursorToY(Yval As Integer)
    If ignoreupdate Then Exit Sub

    Dim Y As Single
    Y = Int(((Yval - 1) / 1024) * (picmap.height - 1))

    Call PlacePointerY(Y)

End Sub


Private Sub BuildRegionPreview()
'          Dim i As Integer
'          Dim j As Integer
    
    SetBkColor picmap.hDC, Region.color
    
    StretchBlt picmap.hDC, 0, 0, picmap.width, picmap.height, Region.disphDC, 0, 0, 1024, 1024, vbSrcCopy
        
'            Region.disp.hDC
'10        For j = 0 To 1023 Step 8
'20            For i = 0 To 1023 Step 8
'30                If Region.bitfield.value(i, j) Then
'40                    Call SetPixel(picmap.hDC, i \ 8, j \ 8, Region.color)
'50                End If
'60            Next
'70        Next

End Sub

Private Sub txtAutowarpY_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAutowarpY.Text) Then
        txtAutowarpY.Text = 0
    Else
        txtAutowarpY.Text = val(txtAutowarpY.Text)
    End If
End Sub


Sub UpdatePreview()
    SetStretchBltMode picmap.hDC, HALFTONE
    'Call StretchBlt(picmap.hdc, 0, 0, picmap.width, picmap.height, parent.pic1024.hdc, 0, 0, 1024, 1024, vbSrcCopy)
    Call parent.cpic1024.stretchToDC(picmap.hDC, 0, 0, picmap.width, picmap.height, 0, 0, 1024, 1024, vbSrcCopy)
    
    Call BuildRegionPreview

    picmap.Refresh
End Sub

Private Sub txtPython_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtPython.tooltiptext = txtPython.Text
End Sub
