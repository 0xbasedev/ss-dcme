VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTextToMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text to Map"
   ClientHeight    =   3195
   ClientLeft      =   -60
   ClientTop       =   210
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox picpreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2280
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   1935
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   6
   End
   Begin VB.CommandButton cmdChangeFont 
      Caption         =   "Change Font"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txttext 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "Enter Text Here"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblstats 
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmTextToMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'width & height of the text
Dim w As Integer
Dim h As Integer

'holds the tile to use
Private parent As frmMain

Public Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Private Sub cmdCancel_Click()
      'Cancels the form
10        Unload Me
End Sub

Private Sub cmdChangeFont_Click()
      'Shows the font dialog
10        On Error GoTo errorh
20        cd.flags = cdlCFScreenFonts
30        cd.ShowFont

          'set the font settings
40        picPreview.FontBold = cd.FontBold
50        picPreview.FontItalic = cd.FontItalic
60        picPreview.FontName = cd.FontName
70        picPreview.FontSize = cd.FontSize
80        picPreview.FontStrikethru = cd.FontStrikethru
90        picPreview.FontUnderline = cd.FontUnderline

          'show the font settings in the label
100       Call ShowStats

          'Updates the preview
110       Call UpdateTextPreview
120       Exit Sub
errorh:
130       If Err = cdlCancel Then
140           Exit Sub
150       End If
End Sub

Private Sub cmdGo_Click()
      'Creates an array, from the black and white pixels in the picture
      'and pass it on to the general form

10        On Error GoTo cmdGo_Click_Error

          'create the array
          Dim Text() As Integer
20        ReDim Text(w, h) As Integer

          Dim i As Integer
          Dim j As Integer
          
          Dim TileToUse As Integer
          
30        If parent.tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
40            TileToUse = parent.tileset.selection(vbLeftButton).tilenr
50        Else
60            TileToUse = 1
70        End If
          
          'check every pixel and if its dark then use the tiletouse
          'else keep it blank
80        For j = 0 To h
90            For i = 0 To w
                  Dim c As Long
100               c = GetPixel(picPreview.hDC, i, j)
110               If GetRED(c) < 128 And GetGREEN(c) < 128 And GetBLUE(c) < 128 Then
120                   Text(i, j) = TileToUse
130               Else
140                   Text(i, j) = 0
150               End If
160           Next
170       Next

          'reenable the general form and executes the ttm
180       Call frmGeneral.ExecuteTextToMap(Text, w + 1, h + 1)

          'disable the go again and unload the form
190       cmdGo.Enabled = False
200       Unload Me

210       On Error GoTo 0
220       Exit Sub

cmdGo_Click_Error:
230       HandleError Err, "frmTextToMap.cmdGo_Click"
End Sub

Private Sub UpdateTextPreview()
      'Draws the preview
      'clear the preview
10        On Error GoTo UpdateTextPreview_Error

20        picPreview.Cls

          'put the correct attributes
30        picPreview.CurrentX = 0
40        picPreview.CurrentY = 0
50        picPreview.ForeColor = vbBlack

          'get the width & height of the text in the textbox
60        w = picPreview.TextWidth(txttext.Text)
70        h = picPreview.TextHeight(txttext.Text)

          ' make sure no out of bounce occurs
80        If w > 1023 Then
90            w = 1023
100       End If
110       If h > 1023 Then
120           h = 1023
130       End If

          ' DO NOT PUT THE PICTUREBOX INTO THE FRAME !!!
          ' as it will be in twips, rather than pixels !
          'resize the picturebox and print the text on it
140       picPreview.width = w
150       picPreview.height = h
160       picPreview.Print txttext.Text

          'enable go as we are good to go
170       cmdGo.Enabled = True

180       On Error GoTo 0
190       Exit Sub

UpdateTextPreview_Error:
200       HandleError Err, "frmTextToMap.UpdateTextPreview"
End Sub

Private Sub Form_Load()
      'disable the general form and the go button
10        Set Me.Icon = frmGeneral.Icon

20        cmdGo.Enabled = False

          'select everything in the textbox
30        txttext.selstart = 0
40        txttext.sellength = Len(txttext)

    cd.FontBold = picPreview.FontBold
    cd.FontName = picPreview.FontName
    cd.FontItalic = picPreview.FontItalic
    cd.FontSize = picPreview.FontSize
        
    
          'update the stats of the current font settings of the
          'picturebox
50        ShowStats

    Call UpdateTextPreview
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      'Cancels the form
10        cmdCancel_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
End Sub

Private Sub txttext_Change()
10        Call UpdateTextPreview
End Sub

Private Sub txttext_Click()
'select everything
'    txttext.SelStart = 0
'    txttext.SelLength = Len(txttext)
End Sub

Sub ShowStats()
      'Show the font settings
10        lblstats.Caption = "Font: " & picPreview.FontName & vbNewLine _
                             & "Size: " & picPreview.FontSize & vbNewLine _
                             & "Bold: " & picPreview.FontBold & vbNewLine _
                             & "Italic: " & picPreview.FontItalic & vbNewLine _
                             & "StrikeThrough: " & picPreview.FontStrikethru & vbNewLine _
                             & "Underline: " & picPreview.FontUnderline
End Sub

