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
10        Call parent.TileText.SetalphaTiles(newalphatile())
End Sub


Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub cmdClear_Click()
10        If MessageBox("Do you really want to clear all assigned tiles?", vbYesNo + vbQuestion, "Clear assigned tiles") = vbYes Then
              Dim i As Integer
20            For i = 0 To 255
30                newalphatile(i) = 0
40            Next
50        End If
60        Call DrawTextPreview

70        pictileset.setfocus

End Sub

Private Sub cmdOK_Click()
10        Call ApplyToMap
20        Unload Me
End Sub

Private Sub Form_Load()
10        Set Me.Icon = frmGeneral.Icon

20        BitBlt pictileset.hDC, 0, 0, pictileset.width, pictileset.height, frmGeneral.cTileset.Pic_Tileset.hDC, 0, 0, vbSrcCopy
30        pictileset.Refresh

End Sub


Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Sub Init()
10        ReDim newalphatile(255) As Integer
20        newalphatile = parent.TileText.GetalphaTiles

30        Call DrawTextPreview
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        Unload Me
End Sub


Sub SetLeftSelection(tilenr As Integer, sizeX As Integer, sizeY As Integer)
      'Sets the left selection of the tileset on the given tilenr
10        If tilenr > 190 Or tilenr + (sizeX - 1) > 190 Or tilenr + ((sizeY - 1) * 19) > 190 Then
20            tilenr = 1
30            sizeX = 1
40            sizeY = 1
50        End If

60        If tilenr = 0 Then tilenr = 190

70        multTileX = sizeX - 1
80        multTileY = sizeY - 1
90        tilesetleft = tilenr

          'move the shape
100       leftsel.Left = ((tilenr - 1) Mod 19) * TILEW
110       leftsel.Top = ((tilenr - 1) \ 19) * TILEW

120       leftsel.width = TILEW * sizeX
130       leftsel.height = TILEW * sizeY

          'Update the tileset preview
140       Call DrawTilePreview(pictileset, leftsel, pictilepreviewleft)
End Sub


Private Sub DrawTilePreview(ByRef srcPic As PictureBox, ByRef srcshape As shape, ByRef destpic As PictureBox)
10        destpic.Cls
20        If srcshape.width > srcshape.height Then
              'Resize considering width

30            If srcshape.width = destpic.width Then
                  'Same size, use bitblt
40                BitBlt destpic.hDC, 0, (destpic.height \ 2) - (srcshape.height \ 2), srcshape.width, srcshape.height, srcPic.hDC, srcshape.Left, srcshape.Top, vbSrcCopy
50            ElseIf srcshape.width < destpic.width Then
                  'Source is smaller, use pixel resize
60                SetStretchBltMode destpic.hDC, COLORONCOLOR
70                StretchBlt destpic.hDC, 0, (destpic.height \ 2) - ((srcshape.height / (srcshape.width / destpic.width)) \ 2), destpic.width, srcshape.height / (srcshape.width / destpic.width), srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
80            Else
                  'Source is larger, use halftone resize
90                SetStretchBltMode destpic.hDC, HALFTONE
100               StretchBlt destpic.hDC, 0, (destpic.height \ 2) - ((srcshape.height / (srcshape.width / destpic.width)) \ 2), destpic.width, srcshape.height / (srcshape.width / destpic.width), srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
110           End If
120       Else
130           If srcshape.height = destpic.height Then
                  'Same size, use bitblt
140               BitBlt destpic.hDC, (destpic.width \ 2) - (srcshape.width \ 2), 0, srcshape.width, srcshape.height, srcPic.hDC, srcshape.Left, srcshape.Top, vbSrcCopy
150           ElseIf srcshape.height < destpic.height Then
                  'Source is smaller, use pixel resize
160               SetStretchBltMode destpic.hDC, COLORONCOLOR
170               StretchBlt destpic.hDC, (destpic.width \ 2) - ((srcshape.width / (srcshape.height / destpic.height)) \ 2), 0, srcshape.width / (srcshape.height / destpic.height), destpic.height, srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
180           Else
                  'Source is larger, use halftone resize
190               SetStretchBltMode destpic.hDC, HALFTONE
200               StretchBlt destpic.hDC, (destpic.width \ 2) - ((srcshape.width / (srcshape.height / destpic.height)) \ 2), 0, srcshape.width / (srcshape.height / destpic.height), destpic.height, srcPic.hDC, srcshape.Left, srcshape.Top, srcshape.width, srcshape.height, vbSrcCopy
210           End If
220       End If
230       destpic.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
10        Set parent = Nothing
End Sub

Private Sub pictileset_KeyDown(KeyCode As Integer, Shift As Integer)

10        If KeyCode = vbKeyLeft Then
20            If IsShift(Shift) Then
30                If tilesetleft Mod 19 <> 1 And multTileX > 1 Then
40                    Call SetLeftSelection(tilesetleft, multTileX, multTileY + 1)
50                End If
60            Else
70                Call SetLeftSelection(tilesetleft - 1, 1, 1)
80            End If


90        ElseIf KeyCode = vbKeyRight Then
100           If IsShift(Shift) Then
110               If (tilesetleft + multTileX) Mod 19 <> 0 Then
120                   Call SetLeftSelection(tilesetleft, multTileX + 2, multTileY + 1)
130               End If
140           Else
150               Call SetLeftSelection(tilesetleft + 1, 1, 1)
160           End If


170       ElseIf KeyCode = vbKeyUp Then
180           If tilesetleft > 19 Then
190               If IsShift(Shift) Then
200                   If multTileY > 1 Then
210                       Call SetLeftSelection(tilesetleft, multTileX + 1, multTileY)
220                   End If
230               Else
240                   Call SetLeftSelection(tilesetleft - 19, 1, 1)
250               End If
260           Else
270               Call SetLeftSelection(171 + tilesetleft, 1, 1)
280           End If


290       ElseIf KeyCode = vbKeyDown Then

300           If IsShift(Shift) Then
310               If (tilesetleft + 19 * multTileY) < 172 Then
320                   Call SetLeftSelection(tilesetleft, multTileX + 1, multTileY + 2)
330               End If
340           ElseIf tilesetleft < 172 Then
350               Call SetLeftSelection(tilesetleft + 19, 1, 1)
360           Else
370               Call SetLeftSelection(tilesetleft - 171, 1, 1)
380           End If


390       ElseIf KeyCode = vbKeyDelete Then
              '''
400       ElseIf KeyCode = vbKeyInsert Then
              '''
410       ElseIf KeyCode = vbKeyHome Then
420           Call SetLeftSelection(tilesetleft - tilesetleft Mod 19 + 1, 1, 1)
430       ElseIf KeyCode = vbKeyEnd Then
440           Call SetLeftSelection(tilesetleft - tilesetleft Mod 19 + 19, 1, 1)
450       ElseIf KeyCode = vbKeyPageUp Then
460           Call SetLeftSelection(tilesetleft Mod 19, 1, 1)
470       ElseIf KeyCode = vbKeyPageDown Then
480           Call SetLeftSelection(tilesetleft Mod 19 + 171, 1, 1)

490       ElseIf KeyCode = vbKeyBack Then
500           Call SetLeftSelection(tilesetleft - 1, 1, 1)
510       End If

End Sub

Private Sub pictileset_KeyPress(KeyAscii As Integer)
          Dim i As Integer
          Dim j As Integer
          Dim nextChar As Integer

          'invalid char, exit
10        If KeyAscii < Asc("!") Then Exit Sub


          'if multiple tiles are selected, and if user typed an alphanumeric character, assign multiple keys at once
20        If (multTileX > 0 Or multTileY > 0) And (IsLcase(KeyAscii) Or IsUcase(KeyAscii) Or IsNumber(KeyAscii)) Then
30            For j = 0 To multTileY
40                For i = 0 To multTileX

                      'check if the next character is the same type as the original one
                      'if not, stop assigning
50                    nextChar = KeyAscii + i + j * (multTileX + 1)
60                    If IsLcase(KeyAscii) And IsLcase(nextChar) Or _
                         IsUcase(KeyAscii) And IsUcase(nextChar) Or _
                         IsNumber(KeyAscii) And IsNumber(nextChar) Then

70                        newalphatile(nextChar) = tilesetleft + i + 19 * j
80                    End If

90                Next
100           Next
110       Else
              'assign tile to character
120           newalphatile(KeyAscii) = tilesetleft

              'select next tile
130           Call SetLeftSelection(tilesetleft + 1, 1, 1)
140       End If


150       Call DrawTextPreview

End Sub

Private Sub pictileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


      'Selects a tile from the tileset
10        If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
              'not in the boundaries of the picture
20            Exit Sub
30        End If

          'indicate the mouse is down
40        SharedVar.MouseDown = True

50        oldtilesetX = (X \ TILEW + 1)
60        oldtilesetY = Y \ TILEW

          'set the selected tile
70        If Button = vbLeftButton Then
80            Call SetLeftSelection(oldtilesetY * 19 + oldtilesetX, 1, 1)
90        End If

End Sub

Private Sub pictileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


      'Still select a tile from the tileset if mousedown, but show the
      'tooltiptext of the different tiles
10        If Not (X > 0 And Y > 0 And X < pictileset.width And Y < pictileset.height) Then
              'not in the boundaries of the picture
20            Exit Sub
30        End If

          ' needed to move the tooltip if it hadn't changed
          'pictileset.ToolTipText = ""

          Dim tilenr As Integer
40        tilenr = (Y \ TILEW) * 19 + (X \ TILEW) + 1

50        pictileset.tooltiptext = TilesetToolTipText(tilenr)

60        If MouseDown Then
              'Call pictileset_MouseDown(Button, Shift, x, y)
              'indicate the mouse is down
70            SharedVar.MouseDown = True

              Dim sizeX As Integer
              Dim sizeY As Integer


              Dim curtilesetX As Integer
              Dim curtilesetY As Integer
80            curtilesetX = (X \ TILEW) + 1
90            curtilesetY = Y \ TILEW

100           sizeX = Abs(curtilesetX - oldtilesetX)
110           sizeY = Abs(curtilesetY - oldtilesetY)

120           If curtilesetX > oldtilesetX Then
130               curtilesetX = oldtilesetX
140           End If
150           If curtilesetY > oldtilesetY Then
160               curtilesetY = oldtilesetY
170           End If

180           If (curtilesetX <= 8 And 8 <= curtilesetX + sizeX) And _
                 (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
190               sizeX = 0
200               sizeY = 0
210           End If
220           If (curtilesetX <= 10 And 10 <= curtilesetX + sizeX) And _
                 (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
230               sizeX = 0
240               sizeY = 0
250           End If
260           If (9 <= curtilesetX + sizeX) And _
                 (curtilesetY <= 13 And 13 <= curtilesetY + sizeY) Then
270               sizeX = 0
280               sizeY = 0
290           End If
300           If (curtilesetX <= 11 And 11 <= curtilesetX + sizeX) And _
                 (curtilesetY <= 11 And 11 <= curtilesetY + sizeY) Then
310               sizeX = 0
320               sizeY = 0
330           End If
340           If tilenr > 256 Then
350               sizeX = 0
360               sizeY = 0
370           End If


380           If sizeX = 0 Then curtilesetX = oldtilesetX
390           If sizeY = 0 Then curtilesetY = oldtilesetY

              'set the selected tile
400           If Button = vbLeftButton Then
410               Call SetLeftSelection(curtilesetY * 19 + curtilesetX, sizeX + 1, sizeY + 1)
420           End If
430       End If

End Sub



Private Sub pictileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'indicate the mouse is up
10        SharedVar.MouseDown = False
End Sub




Private Sub DrawTextPreview()
          Const PRINT_WIDTH = 32
          Const PRINT_HEIGHT = 17

          Dim i As Integer
          Dim printedchars As Integer

          Dim printX As Integer
          Dim printY As Integer

10        picPreview.Cls

20        printedchars = 0

30        For i = 1 To 255

40            If newalphatile(i) <> 0 Or (i >= 97 And i <= 122) Then
50                printX = (printedchars \ (picPreview.height \ PRINT_HEIGHT)) * PRINT_WIDTH
60                picPreview.CurrentX = printX + 1

70                printY = (printedchars Mod (picPreview.height \ PRINT_HEIGHT)) * PRINT_HEIGHT
80                picPreview.CurrentY = printY + 1

90                BitBlt picPreview.hDC, printX + 12, printY, TILEW, TILEW, pictileset.hDC, ((newalphatile(i) - 1) Mod 19) * TILEW, ((newalphatile(i) - 1) \ 19) * TILEW, vbSrcCopy

100               picPreview.Print Chr$(i)
110               printedchars = printedchars + 1
120           End If

130       Next
140       picPreview.Refresh

End Sub


