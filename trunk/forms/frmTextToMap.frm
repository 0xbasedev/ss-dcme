VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
    Set parent = Main
End Sub

Private Sub cmdCancel_Click()
'Cancels the form
    Unload Me
End Sub

Private Sub cmdChangeFont_Click()
'Shows the font dialog
    On Error GoTo errorh
    cd.flags = cdlCFScreenFonts
    cd.ShowFont

    'set the font settings
    picPreview.FontBold = cd.FontBold
    picPreview.FontItalic = cd.FontItalic
    picPreview.FontName = cd.FontName
    picPreview.FontSize = cd.FontSize
    picPreview.FontStrikethru = cd.FontStrikethru
    picPreview.FontUnderline = cd.FontUnderline

    'show the font settings in the label
    Call ShowStats

    'Updates the preview
    Call UpdateTextPreview
    Exit Sub
errorh:
    If Err = cdlCancel Then
        Exit Sub
    End If
End Sub

Private Sub cmdGo_Click()
'Creates an array, from the black and white pixels in the picture
'and pass it on to the general form

    On Error GoTo cmdGo_Click_Error

    'create the array
    Dim Text() As Integer
    ReDim Text(w, h) As Integer

    Dim i As Integer
    Dim j As Integer
    
    Dim TileToUse As Integer
    
    If parent.tileset.selection(vbLeftButton).selectionType = TS_Tiles Then
        TileToUse = parent.tileset.selection(vbLeftButton).tilenr
    Else
        TileToUse = 1
    End If
    
    'check every pixel and if its dark then use the tiletouse
    'else keep it blank
    For j = 0 To h
        For i = 0 To w
            Dim c As Long
            c = GetPixel(picPreview.hDC, i, j)
            If GetRED(c) < 128 And GetGREEN(c) < 128 And GetBLUE(c) < 128 Then
                Text(i, j) = TileToUse
            Else
                Text(i, j) = 0
            End If
        Next
    Next

    'reenable the general form and executes the ttm
    Call frmGeneral.ExecuteTextToMap(Text, w + 1, h + 1)

    'disable the go again and unload the form
    cmdGo.Enabled = False
    Unload Me

    On Error GoTo 0
    Exit Sub

cmdGo_Click_Error:
    HandleError Err, "frmTextToMap.cmdGo_Click"
End Sub

Private Sub UpdateTextPreview()
'Draws the preview
'clear the preview
    On Error GoTo UpdateTextPreview_Error

    picPreview.Cls

    'put the correct attributes
    picPreview.CurrentX = 0
    picPreview.CurrentY = 0
    picPreview.ForeColor = vbBlack

    'get the width & height of the text in the textbox
    w = picPreview.TextWidth(txttext.Text)
    h = picPreview.TextHeight(txttext.Text)

    ' make sure no out of bounce occurs
    If w > 1023 Then
        w = 1023
    End If
    If h > 1023 Then
        h = 1023
    End If

    ' DO NOT PUT THE PICTUREBOX INTO THE FRAME !!!
    ' as it will be in twips, rather than pixels !
    'resize the picturebox and print the text on it
    picPreview.width = w
    picPreview.height = h
    picPreview.Print txttext.Text

    'enable go as we are good to go
    cmdGo.Enabled = True

    On Error GoTo 0
    Exit Sub

UpdateTextPreview_Error:
    HandleError Err, "frmTextToMap.UpdateTextPreview"
End Sub

Private Sub Form_Load()
'disable the general form and the go button
    Set Me.Icon = frmGeneral.Icon

    cmdGo.Enabled = False

    'select everything in the textbox
    txttext.selstart = 0
    txttext.sellength = Len(txttext)

    cd.FontBold = picPreview.FontBold
    cd.FontName = picPreview.FontName
    cd.FontItalic = picPreview.FontItalic
    cd.FontSize = picPreview.FontSize
  
    
    'update the stats of the current font settings of the
    'picturebox
    ShowStats

    Call UpdateTextPreview
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Cancels the form
    cmdCancel_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
End Sub

Private Sub txttext_Change()
    Call UpdateTextPreview
End Sub

Private Sub txttext_Click()
'select everything
'    txttext.SelStart = 0
'    txttext.SelLength = Len(txttext)
End Sub

Sub ShowStats()
'Show the font settings
    lblstats.Caption = "Font: " & picPreview.FontName & vbNewLine _
                       & "Size: " & picPreview.FontSize & vbNewLine _
                       & "Bold: " & picPreview.FontBold & vbNewLine _
                       & "Italic: " & picPreview.FontItalic & vbNewLine _
                       & "StrikeThrough: " & picPreview.FontStrikethru & vbNewLine _
                       & "Underline: " & picPreview.FontUnderline
End Sub

