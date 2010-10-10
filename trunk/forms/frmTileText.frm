VERSION 5.00
Begin VB.Form frmTileText 
   Caption         =   "Assign Tiles To Text"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox picpreview 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.PictureBox pictilepreviewleft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   5640
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   2640
      Width           =   960
   End
   Begin VB.PictureBox pictileset 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   5640
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   1
      Top             =   120
      Width           =   4560
      Begin VB.Shape leftsel 
         BorderColor     =   &H000000FF&
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select a tile, and press the key you want to assign it to."
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   2640
      Width           =   3375
   End
End
Attribute VB_Name = "frmTileText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain
Dim newalphatile() As Integer

Dim oldtilesetX As Integer
Dim oldtilesetY As Integer

Public tilesetleft As Integer
Public multTileX As Integer
Public multTileY As Integer



Private Sub ApplyToMap()
    Call parent.TileText.SetalphaTiles(newalphatile())
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    If MessageBox("Do you really want to clear all assigned tiles?", vbYesNo + vbQuestion, "Clear assigned tiles") = vbYes Then
        Dim i As Integer
        For i = 0 To 255
            newalphatile(i) = 0
        Next
    End If
    Call DrawTextPreview

    pictileset.setfocus

End Sub

Private Sub cmdOK_Click()
    Call ApplyToMap
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmGeneral.Icon

    BitBlt pictileset.hDC, 0, 0, pictileset.width, pictileset.height, frmGeneral.cTileset.Pic_Tileset.hDC, 0, 0, vbSrcCopy
    pictileset.Refresh

End Sub


Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub Init()
    ReDim newalphatile(255) As Integer
    newalphatile = parent.TileText.GetalphaTiles

    Call DrawTextPreview
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub


Sub SetLeftSelection(tilenr As Integer, sizeX As Integer, sizeY As Integer)
'Sets the left selection of the tileset on the given tilenr
    If tilenr > 190 Or tilenr + (sizeX - 1) > 190 Or tilenr + ((sizeY - 1) * 19) > 190 Then
        tilenr = 1
        sizeX = 1
        sizeY = 1
    End If

    If tilenr = 0 Then tilenr = 190

    multTileX = sizeX - 1
    multTileY = sizeY - 1
    tilesetleft = tilenr

    'move the shape
    leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
    leftsel.Top = ((tilenr - 1) \ 19) * TILEW

    leftsel.width = TILEW * sizeX
    leftsel.height = TILEW * sizeY

    'Update the tileset preview
    Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)
End Sub


Private Sub DrawTilePreview(ByRef srcPic As PictureBox, ByRef srcshape As shape, ByRef destpic As PictureBox)
    destpic.Cls
    If srcshape.width > srcshape.height Then
        'Resize considering width

        If srcshape.width = destpic.width Then
            'Same size, use bitblt
            BitBlt destpic.hDC, 0, (destpic.height \ 2) - (srcshape.height \ 2), srcshape.width, srcshape.height, srcPic.hDC, srcshape.Left, srcshape.Top, vbSrcCopy
        ElseIf srcshape.width < destpic.width Then
            'Source is smaller, use pixel resize
            SetStretchBltMode destpic.hDC, COLORONCOLOR
            StretchBlt destpic.hDC, 0, (destpic.height \ 2) - ((srcshape.height / (srcshape.width / destpic.width)) \ 2), destpic.width, srcshape.height / (srcshape.width / destpic.width), srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
        Else
            'Source is larger, use halftone resize
            SetStretchBltMode destpic.hDC, HALFTONE
            StretchBlt destpic.hDC, 0, (destpic.height \ 2) - ((srcshape.height / (srcshape.width / destpic.width)) \ 2), destpic.width, srcshape.height / (srcshape.width / destpic.width), srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
        End If
    Else
        If srcshape.height = destpic.height Then
            'Same size, use bitblt
            BitBlt destpic.hDC, (destpic.width \ 2) - (srcshape.width \ 2), 0, srcshape.width, srcshape.height, srcPic.hDC, srcshape.Left, srcshape.Top, vbSrcCopy
        ElseIf srcshape.height < destpic.height Then
            'Source is smaller, use pixel resize
            SetStretchBltMode destpic.hDC, COLORONCOLOR
            StretchBlt destpic.hDC, (destpic.width \ 2) - ((srcshape.width / (srcshape.height / destpic.height)) \ 2), 0, srcshape.width / (srcshape.height / destpic.height), destpic.height, srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
        Else
            'Source is larger, use halftone resize
            SetStretchBltMode destpic.hDC, HALFTONE
            StretchBlt destpic.hDC, (destpic.width \ 2) - ((srcshape.width / (srcshape.height / destpic.height)) \ 2), 0, srcshape.width / (srcshape.height / destpic.height), destpic.height, srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
        End If
    End If
    destpic.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set parent = Nothing
End Sub

Private Sub pictileset_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then
        If IsShift(Shift) Then
            If tilesetleft Mod 19 <> 1 And multTileX > 1 Then
                Call SetLeftSelection(tilesetleft, multTileX, multTileY + 1)
            End If
        Else
            Call SetLeftSelection(tilesetleft - 1, 1, 1)
        End If


    ElseIf KeyCode = vbKeyRight Then
        If IsShift(Shift) Then
            If (tilesetleft + multTileX) Mod 19 <> 0 Then
                Call SetLeftSelection(tilesetleft, multTileX + 2, multTileY + 1)
            End If
        Else
            Call SetLeftSelection(tilesetleft + 1, 1, 1)
        End If


    ElseIf KeyCode = vbKeyUp Then
        If tilesetleft > 19 Then
            If IsShift(Shift) Then
                If multTileY > 1 Then
                    Call SetLeftSelection(tilesetleft, multTileX + 1, multTileY)
                End If
            Else
                Call SetLeftSelection(tilesetleft - 19, 1, 1)
            End If
        Else
            Call SetLeftSelection(171 + tilesetleft, 1, 1)
        End If


    ElseIf KeyCode = vbKeyDown Then

        If IsShift(Shift) Then
            If (tilesetleft + 19 * multTileY) < 172 Then
                Call SetLeftSelection(tilesetleft, multTileX + 1, multTileY + 2)
            End If
        ElseIf tilesetleft < 172 Then
            Call SetLeftSelection(tilesetleft + 19, 1, 1)
        Else
            Call SetLeftSelection(tilesetleft - 171, 1, 1)
        End If


    ElseIf KeyCode = vbKeyDelete Then
        '''
    ElseIf KeyCode = vbKeyInsert Then
        '''
    ElseIf KeyCode = vbKeyHome Then
        Call SetLeftSelection(tilesetleft - tilesetleft Mod 19 + 1, 1, 1)
    ElseIf KeyCode = vbKeyEnd Then
        Call SetLeftSelection(tilesetleft - tilesetleft Mod 19 + 19, 1, 1)
    ElseIf KeyCode = vbKeyPageUp Then
        Call SetLeftSelection(tilesetleft Mod 19, 1, 1)
    ElseIf KeyCode = vbKeyPageDown Then
        Call SetLeftSelection(tilesetleft Mod 19 + 171, 1, 1)

    ElseIf KeyCode = vbKeyBack Then
        Call SetLeftSelection(tilesetleft - 1, 1, 1)
    End If

End Sub

Private Sub pictileset_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim nextChar As Integer

    'invalid char, exit
    If KeyAscii < Asc("!") Then Exit Sub


    'if multiple tiles are selected, and if user typed an alphanumeric character, assign multiple keys at once
    If (multTileX > 0 Or multTileY > 0) And (IsLcase(KeyAscii) Or IsUcase(KeyAscii) Or IsNumber(KeyAscii)) Then
        For j = 0 To multTileY
            For i = 0 To multTileX

                'check if the next character is the same type as the original one
                'if not, stop assigning
                nextChar = KeyAscii + i + j * (multTileX + 1)
                If IsLcase(KeyAscii) And IsLcase(nextChar) Or _
                   IsUcase(KeyAscii) And IsUcase(nextChar) Or _
                   IsNumber(KeyAscii) And IsNumber(nextChar) Then

                    newalphatile(nextChar) = tilesetleft + i + 19 * j
                End If

            Next
        Next
    Else
        'assign tile to character
        newalphatile(KeyAscii) = tilesetleft

        'select next tile
        Call SetLeftSelection(tilesetleft + 1, 1, 1)
    End If


    Call DrawTextPreview

End Sub

Private Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


'Selects a tile from the tileset
    If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
        'not in the boundaries of the picture
        Exit Sub
    End If

    'indicate the mouse is down
    SharedVar.MouseDown = True

    oldtilesetX = (X \ TILEW + 1)
    oldtilesetY = Y \ TILEW

    'set the selected tile
    If Button = vbLeftButton Then
        Call SetLeftSelection(oldtilesetY * 19 + oldtilesetX, 1, 1)
    End If

End Sub

Private Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


'Still select a tile from the tileset if mousedown, but show the
'tooltiptext of the different tiles
    If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
        'not in the boundaries of the picture
        Exit Sub
    End If

    ' needed to move the tooltip if it hadn't changed
    'pictileset.ToolTipText = ""

    Dim tilenr As Integer
    tilenr = (Y \ TILEW) * 19 + (X \ TILEW) + 1

    pictileset.tooltiptext = TilesetToolTipText(tilenr)

    If MouseDown Then
        'Call pictileset_MouseDown(Button, Shift, x, y)
        'indicate the mouse is down
        SharedVar.MouseDown = True

        Dim sizeX As Integer
        Dim sizeY As Integer


        Dim curtilesetX As Integer
        Dim curtilesetY As Integer
        curtilesetX = (X \ TILEW) + 1
        curtilesetY = Y \ TILEW

        sizeX = Abs(curtilesetX - oldtilesetX)
        sizeY = Abs(curtilesetY - oldtilesetY)

        If curtilesetX > oldtilesetX Then
            curtilesetX = oldtilesetX
        End If
        If curtilesetY > oldtilesetY Then
            curtilesetY = oldtilesetY
        End If

        If (curtilesetX <= 8 And 8 <= curtilesetX + sizeX) And _
           (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
            sizeX = 0
            sizeY = 0
        End If
        If (curtilesetX <= 10 And 10 <= curtilesetX + sizeX) And _
           (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
            sizeX = 0
            sizeY = 0
        End If
        If (9 <= curtilesetX + sizeX) And _
           (curtilesetY <= 13 And 13 <= curtilesetY + sizeY) Then
            sizeX = 0
            sizeY = 0
        End If
        If (curtilesetX <= 11 And 11 <= curtilesetX + sizeX) And _
           (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
            sizeX = 0
            sizeY = 0
        End If
        If tilenr > 256 Then
            sizeX = 0
            sizeY = 0
        End If


        If sizeX = 0 Then curtilesetX = oldtilesetX
        If sizeY = 0 Then curtilesetY = oldtilesetY

        'set the selected tile
        If Button = vbLeftButton Then
            Call SetLeftSelection(curtilesetY * 19 + curtilesetX, sizeX + 1, sizeY + 1)
        End If
    End If

End Sub



Private Sub pictileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'indicate the mouse is up
    SharedVar.MouseDown = False
End Sub




Private Sub DrawTextPreview()
    Const PRINT_WIDTH = 32
    Const PRINT_HEIGHT = 17

    Dim i As Integer
    Dim printedchars As Integer

    Dim printX As Integer
    Dim printY As Integer

    picPreview.Cls

    printedchars = 0

    For i = 1 To 255

        If newalphatile(i) <> 0 Or (i >= 97 And i <= 122) Then
            printX = (printedchars \ (picPreview.height \ PRINT_HEIGHT)) * PRINT_WIDTH
            picPreview.CurrentX = printX + 1

            printY = (printedchars Mod (picPreview.height \ PRINT_HEIGHT)) * PRINT_HEIGHT
            picPreview.CurrentY = printY + 1

            BitBlt picPreview.hDC, printX + 12, printY, TILEW, TILEW, pictileset.hDC, ((newalphatile(i) - 1) Mod 19) * TILEW, ((newalphatile(i) - 1) \ 19) * TILEW, vbSrcCopy

            picPreview.Print Chr$(i)
            printedchars = printedchars + 1
        End If

    Next
    picPreview.Refresh

End Sub


