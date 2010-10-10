VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CfgEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings Editor"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6300
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "File: example.cfg"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1920
      Left            =   6720
      TabIndex        =   6
      Top             =   720
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3387
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
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ship"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Index           =   1
      Left            =   240
      ScaleHeight     =   5655
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      Begin VB.CommandButton Command3 
         Caption         =   "Error Check"
         Height          =   375
         Left            =   6360
         TabIndex        =   10
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export To Continuum"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply To CFG"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   4
         Top             =   4080
         Width           =   8895
         Begin VB.ListBox List1 
            Height          =   1230
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ship Comparisons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   6360
         TabIndex        =   3
         Top             =   0
         Width           =   2535
         Begin VB.Frame Frame3 
            Caption         =   "Current Setting"
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   2160
            Width           =   2295
         End
      End
      Begin MSComctlLib.ListView Settings 
         Height          =   4080
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   7197
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
            Text            =   "Setting"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tbsCfgEditor 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10821
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Warbird"
            Key             =   "wb"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Javelin"
            Key             =   "jav"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Spider"
            Key             =   "spid"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Leviathan"
            Key             =   "levi"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Terrier"
            Key             =   "ter"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Weasel"
            Key             =   "weasel"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lancaster"
            Key             =   "lanc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Shark"
            Key             =   "shark"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other"
            Key             =   "other"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "CfgEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tbsCfgEditor_Click()
If tbsCfgEditor.SelectedItem.Caption = "Other" Then MsgBox "Hi", vbCritical
End Sub
