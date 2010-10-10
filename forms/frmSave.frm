VERSION 5.00
Begin VB.Form frmSave 
   Caption         =   "Save"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   3390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTileset 
      Caption         =   "Save Tileset"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "If unchecked, the default tileset will be used for this map. Note that the tileset is also required for all eLVL data."
      Top             =   840
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox chkLVZ 
      Caption         =   "Save LVZ Files"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Saves all LVZ files associated with your map."
      Top             =   480
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame frmeLVL 
      Caption         =   "eLVL data"
      Height          =   2055
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   3255
      Begin VB.CheckBox chkelvlLVZ 
         Caption         =   "Save LVZ Paths"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Allows DCME to re-open LVZ files automatically with the map."
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkelvlTT 
         Caption         =   "Save Text-Tiles Definitions"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Text-Tiles definitions will be saved into the map file."
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkelvlWT 
         Caption         =   "Save Walltiles Definitions"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Walltiles definitions will be saved into the map file."
         Top             =   960
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkelvlREGN 
         Caption         =   "Save Regions"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "ASSS regions will be saved."
         Top             =   600
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkelvlATTR 
         Caption         =   "Save Map Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Textual attributes will be included with the map."
         Top             =   240
         Value           =   1  'Checked
         Width           =   3015
      End
   End
   Begin VB.CheckBox chkNoExtraTiles 
      Caption         =   "Save as SSME compatible"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "All extra tiles will be removed from the map. For full SSME compatibility, a 8bit tileset is also required."
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkTileset_Click()

10        chkelvlATTR.Enabled = chkTileset.value = vbChecked
20        chkelvlREGN.Enabled = chkTileset.value = vbChecked
30        chkelvlWT.Enabled = chkTileset.value = vbChecked
40        chkelvlTT.Enabled = chkTileset.value = vbChecked
50        chkelvlLVZ.Enabled = chkTileset.value = vbChecked

End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub cmdSave_Click()
          Dim flags As saveFlags
          
10        flags = SFdefault
          
20        FlagAdd flags, chkNoExtraTiles.value = vbUnchecked, SFsaveExtraTiles
          
30        FlagAdd flags, chkLVZ.value = vbChecked, SFsaveLVZ
          
          
      '    If chkNoExtraTiles.Value = vbChecked Then
      '        flags = flags And Not SFsaveExtraTiles
      '    End If
          
      '    If chkLVZ.Value = vbChecked Then
      '        flags = flags And SFsaveLVZ
      '    Else
      '        flags = flags And Not SFsaveLVZ
      '    End If
          
40        If chkTileset.value = vbChecked Then
              'ELVL needs tileset
50            FlagAdd flags, True, SFsaveTileset
          
      '        flags = flags And SFsaveTileset
              
60            FlagAdd flags, chkelvlATTR.value = vbChecked, SFsaveELVLattr
70            FlagAdd flags, chkelvlREGN.value = vbChecked, SFsaveELVLregn
80            FlagAdd flags, chkelvlWT.value = vbChecked, SFsaveELVLdcwt
90            FlagAdd flags, chkelvlTT.value = vbChecked, SFsaveELVLdctt
100           FlagAdd flags, chkelvlLVZ.value = vbChecked, SFsaveELVLdclv
          
          
      '        If chkelvlATTR.Value = vbChecked Then
      '            flags = flags And SFsaveELVLattr
      '        Else
      '            flags = flags And Not SFsaveELVLattr
      '        End If
              
      '        If chkelvlREGN.Value = vbChecked Then
      '            flags = flags And SFsaveELVLregn
      '        Else
      '            flags = flags And Not SFsaveELVLregn
      '        End If
      '
      '        If chkelvlWT.Value = vbChecked Then
      '            flags = flags And SFsaveELVLdcwt
      '        Else
      '            flags = flags And Not SFsaveELVLdcwt
      '        End If
      '
      '        If chkelvlTT.Value = vbChecked Then
      '            flags = flags And SFsaveELVLdctt
      '        Else
      '            flags = flags And Not SFsaveELVLdctt
      '        End If
      '
      '        If chkelvlLVZ.Value = vbChecked Then
      '            flags = flags And SFsaveELVLdclv
      '        Else
      '            flags = flags And Not SFsaveELVLdclv
      '        End If
110       Else
120           FlagAdd flags, False, SFsaveTileset
130           FlagAdd flags, False, SFsaveELVL
              
      '        flags = flags And Not SFsaveTileset
      '        flags = flags And Not SFsaveELVL
140       End If
          

          
150       Unload Me
160       Call frmGeneral.SaveMap(True, flags)

End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Unload Me
End Sub

