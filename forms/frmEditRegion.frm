VERSION 5.00
Begin VB.Form frmEditRegion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Region"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPython 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmEditRegion.frx":0000
      Top             =   480
      Width           =   615
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
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Jump To"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
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
      Top             =   1320
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
         Y1              =   0
         Y2              =   129
      End
   End
   Begin VB.TextBox txtArena 
      Height          =   285
      Left            =   480
      MaxLength       =   16
      TabIndex        =   13
      ToolTipText     =   "Name of the arena to warp to. Leave blank for current arena."
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtAutowarpY 
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "512"
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtAutowarpX 
      Height          =   285
      Left            =   480
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "512"
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CheckBox chkAutowarp 
      Caption         =   "Autowarp"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
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
      Left            =   0
      TabIndex        =   12
      ToolTipText     =   "Name of the arena to warp to. Leave blank for current arena"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblY 
      Caption         =   "Y:"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   2670
      Width           =   255
   End
   Begin VB.Label lblX 
      Caption         =   "X:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Set to 1-1024 for autowarp coordinates, Set to 0 for arena default, Set to -1 for current player position"
      Top             =   2670
      Width           =   255
   End
End
Attribute VB_Name = "frmEditRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WIDTH1 = 4155    'width of form with autowarp Off
Const WIDTH2 = 4155    'width of form with autowarp On
Const WIDTH3 = 4155    'width of form when viewing the minimap to select a coordinate

Const HEIGHT1 = 2940
Const HEIGHT2 = 3735
Const HEIGHT3 = 3735
Dim parent As frmMain
Dim Region As Region
Dim rIdx As Integer
Dim ignoreupdate As Boolean

Sub setParent(Main As frmMain, regionindex As Integer)
10        Set parent = Main
20        rIdx = regionindex

30        Call LoadRegion
40        Call UpdatePreview
End Sub

Private Sub chkAutowarp_Click()
10        CheckAutowarpType
End Sub

Private Sub cmdEditPython_Click()
10        frmPython.show vbModal, frmGeneral
20        Region.pythonCode = frmGeneral.GetActiveRegionPythonCode
30        If Region.pythonCode <> "" Then
40            txtPython = "..."
50        Else
60            txtPython = "<None>"
70        End If
End Sub

Private Sub cmdGoto_Click()
          Dim tileX As Integer, tileY As Integer

10        tileX = val(txtAutowarpX.Text)
20        tileY = val(txtAutowarpY.Text)

30        If tileX = 0 Or tileY = 0 Then
40            MessageBox "Invalid jump coordinates. 0 is used for arena's default warp coordinate.", vbOKOnly + vbExclamation, "Invalid coordinate"
50        ElseIf tileX = -1 Or tileY = -1 Then
60            MessageBox "Invalid jump coordinates. -1 is used for player's current position.", vbOKOnly + vbExclamation, "Invalid coordinate"
70        Else
80            Call parent.SetFocusAt(tileX - 1, tileY - 1, parent.picPreview.width \ 2, parent.picPreview.height \ 2, True)
90        End If
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon

20        Me.width = WIDTH3
30        Me.height = HEIGHT3

40        Me.Left = frmGeneral.Left + frmGeneral.llRegionList.Left + GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelY
50        Me.Top = frmGeneral.Top + frmGeneral.tlbToolOptions.Top + frmGeneral.tlbToolOptions.height + (GetSystemMetrics(SM_CYMENU) + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFRAME)) * Screen.TwipsPerPixelY
    
    Call ShowQuickMap
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If picmap.visible Then
'        Call HideQuickMap
'    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Call SaveRegion
20        Unload Me
End Sub

Private Sub LoadRegion()
10        Set Region = parent.Regions.getRegion(rIdx)

20        If Region.isBase Then
30            chkIsBase.value = vbChecked
40        Else
50            chkIsBase.value = vbUnchecked
60        End If

70        If Region.isNoAntiwarp Then
80            chkNoAntiWarp.value = vbChecked
90        Else
100           chkNoAntiWarp.value = vbUnchecked
110       End If

120       If Region.isNoWeapon Then
130           chkNoWeapons.value = vbChecked
140       Else
150           chkNoWeapons.value = vbUnchecked
160       End If

170       If Region.isNoFlagDrop Then
180           chkNoFlagDrops.value = vbChecked
190       Else
200           chkNoFlagDrops.value = vbUnchecked
210       End If

220       If Region.isAutoWarp Then
230           chkAutowarp.value = vbChecked
240       Else
250           chkAutowarp.value = vbUnchecked
260       End If

270       If Region.pythonCode <> "" Then
280           txtPython = "..."
290       Else
300           txtPython = "<None>"
310       End If
320       txtName.Text = Region.name
          
330       txtAutowarpX.Text = Region.autowarpX
340       txtAutowarpY.Text = Region.autowarpY

350       txtArena.Text = Region.autowarpArena

360       Call CheckAutowarpType

End Sub

Private Sub SaveRegion()
10        If txtName.Text <> "" Then
20            Region.name = txtName.Text
30        End If
40        Region.autowarpArena = txtArena.Text
50        Region.autowarpX = val(txtAutowarpX.Text)
60        Region.autowarpY = val(txtAutowarpY.Text)
70        Region.isAutoWarp = (chkAutowarp.value = vbChecked)
80        Region.isBase = (chkIsBase.value = vbChecked)
90        Region.isNoAntiwarp = (chkNoAntiWarp.value = vbChecked)
100       Region.isNoFlagDrop = (chkNoFlagDrops.value = vbChecked)
110       Region.isNoWeapon = (chkNoWeapons.value = vbChecked)
          'the python code was already saved in region
120       Region.pythonCode = frmGeneral.GetActiveRegionPythonCode
          
      '    Call parent.Regions.setRegion(Region, rIdx)
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
20        Set Region = Nothing
End Sub

Private Sub txtAutowarpX_Change()
          Dim newX As Integer
10        Call removeDisallowedCharacters(txtAutowarpX, -1, 1024)

20        newX = val(txtAutowarpX.Text)
30        If newX = 0 Or newX = -1 Then
40            newX = 512
50        End If
60        Call MoveCursorToX(newX)
End Sub

Private Sub txtAutowarpX_Validate(Cancel As Boolean)
10        If Not IsNumeric(txtAutowarpX.Text) Then
20            txtAutowarpX.Text = 0
30        Else
40            txtAutowarpX.Text = val(txtAutowarpX.Text)
50        End If
End Sub

Private Sub txtAutowarpY_Change()
          Dim newY As Integer
10        Call removeDisallowedCharacters(txtAutowarpY, -1, 1024)

20        newY = val(txtAutowarpY.Text)
30        If newY = 0 Or newY = -1 Then
40            newY = 512
50        End If
60        Call MoveCursorToY(newY)
End Sub



Private Sub ShowQuickMap()
10        cmdGoto.visible = True
20        picmap.visible = True
30        Me.width = WIDTH3
40        Me.height = HEIGHT3

End Sub

Private Sub HideQuickMap()
10        cmdGoto.visible = False
20        picmap.visible = False
30        Me.width = WIDTH2
40        Me.height = HEIGHT2
End Sub



Private Sub CheckAutowarpType()
10        If chkAutowarp.value = vbChecked Then
20            txtAutowarpX.Enabled = True
30            txtAutowarpY.Enabled = True
40            lblX.Enabled = True
50            lblY.Enabled = True
60            txtArena.Enabled = True
70            lblArena.Enabled = True

'80            Me.width = WIDTH2
'90            Me.height = HEIGHT2
'100           Call ShowQuickMap


110       Else
120           txtAutowarpX.Enabled = False
130           txtAutowarpY.Enabled = False
140           lblX.Enabled = False
150           lblY.Enabled = False
160           txtArena.Enabled = False
170           lblArena.Enabled = False

'180           Call HideQuickMap

'190           Me.width = WIDTH1
'200           Me.height = HEIGHT1
210       End If
End Sub


Private Sub picmap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If X <= 0 Then X = 0
20        If X >= picmap.ScaleWidth - 1 Then X = picmap.ScaleWidth - 1
30        If Y <= 0 Then Y = 0
40        If Y >= picmap.ScaleHeight - 1 Then Y = picmap.ScaleHeight - 1

50        Call PlacePointer(X, Y)
60        Call UpdateValues
End Sub

Private Sub picmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If Button Then
20            Call picmap_MouseDown(Button, Shift, X, Y)
30        End If

End Sub

Private Sub PlacePointer(X As Single, Y As Single)
10        Call PlacePointerX(X)
20        Call PlacePointerY(Y)
End Sub

Private Sub PlacePointerX(X As Single)
10        cursor.Left = X - 3

20        cursorLine1.x1 = X
30        cursorLine1.x2 = X

40        cursorLine2.x1 = X - 5
50        cursorLine2.x2 = X + 5
End Sub

Private Sub PlacePointerY(Y As Single)
10        cursor.Top = Y - 3

20        cursorLine1.y1 = Y - 5
30        cursorLine1.y2 = Y + 5

40        cursorLine2.y1 = Y
50        cursorLine2.y2 = Y
End Sub


Private Sub UpdateValues()
10        ignoreupdate = True
20        txtAutowarpX.Text = Int(((cursor.Left + 3) / (picmap.width - 1)) * 1023) + 1
30        txtAutowarpY.Text = Int(((cursor.Top + 3) / (picmap.height - 1)) * 1023) + 1
40        ignoreupdate = False
End Sub

Private Sub MoveCursorToX(Xval As Integer)
10        If ignoreupdate Then Exit Sub

          Dim X As Single
20        X = Int(((Xval - 1) / 1024) * (picmap.width - 1))

30        Call PlacePointerX(X)
End Sub

Private Sub MoveCursorToY(Yval As Integer)
10        If ignoreupdate Then Exit Sub

          Dim Y As Single
20        Y = Int(((Yval - 1) / 1024) * (picmap.height - 1))

30        Call PlacePointerY(Y)

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
10        If Not IsNumeric(txtAutowarpY.Text) Then
20            txtAutowarpY.Text = 0
30        Else
40            txtAutowarpY.Text = val(txtAutowarpY.Text)
50        End If
End Sub


Sub UpdatePreview()
10        SetStretchBltMode picmap.hDC, HALFTONE
          'Call StretchBlt(picmap.hdc, 0, 0, picmap.width, picmap.height, parent.pic1024.hdc, 0, 0, 1024, 1024, vbSrcCopy)
20        Call parent.cpic1024.stretchToDC(picmap.hDC, 0, 0, picmap.width, picmap.height, 0, 0, 1024, 1024, vbSrcCopy)
          
30        Call BuildRegionPreview

40        picmap.Refresh
End Sub

Private Sub txtPython_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        txtPython.tooltiptext = txtPython.Text
End Sub
