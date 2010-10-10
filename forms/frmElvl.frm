VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmElvl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eLVL Properties"
   ClientHeight    =   5445
   ClientLeft      =   -5805
   ClientTop       =   4995
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRegions 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   0
      Left            =   6840
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   0
      Left            =   7320
      ScaleHeight     =   4335
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtPropertyEdit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddProperty 
         Caption         =   "Add Property"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemoveProperty 
         Caption         =   "Remove Property"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddStandardProperties 
         Caption         =   "Add Standard Properties"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ListView lstProperties 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Attribute"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip tbsElvl 
      Height          =   4815
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Attributes"
            Key             =   "Attributes"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmElvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim propEdit As Boolean
Dim curlstItem As ListItem
Dim editVal As Boolean


Dim curX As Single
Dim curY As Single

Dim parent As frmMain

'contains the regions where we are working on, if we cancel we discard
'them , if we press ok we replace the old map ones with these ones



Private Sub cmdAddProperty_Click()

    Dim idx As Integer
    If lstProperties.SelectedItem Is Nothing Then
        idx = 1
    Else
        idx = (lstProperties.SelectedItem.Index + 1)
    End If
    Call lstProperties.ListItems.add(idx, , "new Property")
    lstProperties.ListItems.item(idx).SubItems(1) = "new value"

End Sub

Private Sub AddStandardProperties()
    Dim idx As Integer
    idx = lstProperties.ListItems.count

    If Not itemExistsInList(lstProperties, "Name", False) Then
        Call lstProperties.ListItems.add(, "Name", "Name")
        lstProperties.ListItems.item(lstProperties.ListItems.count).SubItems(1) = GetFileNameWithoutExtension(parent.Caption)
    End If

    If Not itemExistsInList(lstProperties, "Version", False) Then
        Call lstProperties.ListItems.add(, "Version", "Version")
        lstProperties.ListItems.item(lstProperties.ListItems.count).SubItems(1) = "1.0"
    End If

    If Not itemExistsInList(lstProperties, "Zone", False) Then
        Call lstProperties.ListItems.add(, "Zone", "Zone")
    End If

    If Not itemExistsInList(lstProperties, "MapCreator", False) Then
        Call lstProperties.ListItems.add(, "MapCreator", "MapCreator")
        lstProperties.ListItems.item(lstProperties.ListItems.count).SubItems(1) = GetUserName
    End If

    If Not itemExistsInList(lstProperties, "Program", False) Then
        Call lstProperties.ListItems.add(, "Program", "Program")
        lstProperties.ListItems.item(lstProperties.ListItems.count).SubItems(1) = "DCME " & App.Major & "." & App.Minor & "." & App.Revision
    End If

    If Not itemExistsInList(lstProperties, "TilesetCreator", False) Then
        Call lstProperties.ListItems.add(, "TilesetCreator", "TilesetCreator")
    End If
End Sub






Private Sub cmdRemoveProperty_Click()
    If lstProperties.SelectedItem Is Nothing Then
    Else
        Call lstProperties.ListItems.Remove(lstProperties.SelectedItem.Index)
    End If
End Sub








Private Sub lstProperties_Click()
    ClosePropertiesEdit
End Sub

Private Sub lstProperties_DblClick()
    If propEdit Then
        ClosePropertiesEdit
    End If


    Set curlstItem = lstProperties.HitTest(curX, curY)
    If curlstItem Is Nothing Then Exit Sub

    propEdit = True

    If curX > lstProperties.ColumnHeaders(1).width Then
        'we hovered over the value
        editVal = True
    Else
        editVal = False
    End If

    If propEdit Then
        If editVal Then
            txtPropertyEdit.Left = lstProperties.Left + curlstItem.Left + lstProperties.ColumnHeaders(1).width
            txtPropertyEdit.width = lstProperties.ColumnHeaders(2).width
            txtPropertyEdit.Top = lstProperties.Top + curlstItem.Top
            txtPropertyEdit.height = curlstItem.height

            txtPropertyEdit.Text = curlstItem.SubItems(1)
        Else
            txtPropertyEdit.Left = lstProperties.Left + curlstItem.Left
            txtPropertyEdit.width = lstProperties.ColumnHeaders(1).width
            txtPropertyEdit.Top = lstProperties.Top + curlstItem.Top
            txtPropertyEdit.height = curlstItem.height

            txtPropertyEdit.Text = curlstItem.Text
        End If

        txtPropertyEdit.visible = True

        On Error Resume Next
        txtPropertyEdit.setfocus
    Else
        If editVal Then
            curlstItem.SubItems(1) = txtPropertyEdit.Text
        Else
            curlstItem.Text = txtPropertyEdit.Text
            curlstItem.Key = txtPropertyEdit.Text
        End If
        txtPropertyEdit.visible = False
    End If

End Sub

Private Sub lstProperties_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    curX = X
    curY = Y
End Sub







Private Sub txtPropertyEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ClosePropertiesEdit
    End If
End Sub


Sub ClosePropertiesEdit()
    If propEdit Then
        If editVal Then
            curlstItem.SubItems(1) = txtPropertyEdit.Text
        Else
            curlstItem.Text = txtPropertyEdit.Text
            curlstItem.Key = txtPropertyEdit.Text
        End If
        txtPropertyEdit.visible = False

        propEdit = False
    End If
End Sub

Private Function itemExistsInList(l As ListView, str As String, casesensitive As Boolean) As Boolean
    Dim i As Integer
    For i = 1 To l.ListItems.count
        If l.ListItems(i).Text = str Or _
           (Not casesensitive And LCase(l.ListItems(i).Text) = LCase(str)) Then
            itemExistsInList = True
            Exit Function
        End If
    Next

    itemExistsInList = False
End Function





Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    'divide list in 2 equal halves
    lstProperties.ColumnHeaders(1).width = (lstProperties.width - 60) \ 2
    lstProperties.ColumnHeaders(2).width = (lstProperties.width - 60) \ 2

    lstProperties.ListItems.Clear
    Call parent.eLVL.getAttributeList(lstProperties.ListItems)


    Call AddStandardProperties


    Call tbsElvl_Click
End Sub

Sub setParent(map As frmMain)
10        Set parent = map
End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub cmdOK_Click()
'apply the property list
    Call parent.eLVL.setAttributeList(lstProperties.ListItems)

'    Call parent.Regions.BuildRegionTiles

    parent.mapchanged = True

    Unload Me
End Sub

Private Sub tbsElvl_Click()

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsElvl.Tabs.count - 1
        If i = tbsElvl.SelectedItem.Index - 1 Then
            picTab(i).Left = 15
            picTab(i).Top = 32
            picTab(i).visible = True
        Else
            picTab(i).visible = False
        End If
    Next

End Sub



'
'Private Sub setlstToolTipText(Key As String, tooltiptext As String)
'    lstRegionProperties.ListItems(Key).tooltiptext = tooltiptext
'    lstRegionProperties.ListItems(Key).ListSubItems(1).tooltiptext = tooltiptext
'End Sub





Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
End Sub


