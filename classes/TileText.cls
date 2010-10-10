VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim parent As frmMain

Dim dragStart As Integer
Dim dragEnd As Integer
Dim dragging As Boolean

Dim Active As Boolean
Dim Boundaries As area

Dim alphaTile() As Integer

Dim insert As Boolean

Dim lastkey As Integer

Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Function GetalphaTiles() As Integer()
10        GetalphaTiles = alphaTile
End Function

Function GetalphaTile(char As Integer) As Integer
10        GetalphaTile = alphaTile(char)
End Function

Sub SetalphaTiles(newalphatile() As Integer)
10        alphaTile() = newalphatile()
End Sub

Sub SetalphaTile(char As Integer, tile As Integer)
10        alphaTile(char) = tile
End Sub

Function isCharacterCustom(char As Integer) As Boolean
          'Returns true if the tile defined for the specified character is different from default
10        If char >= Asc("a") And char <= Asc("z") Then
20            isCharacterCustom = (alphaTile(char) <> char - Asc("a") + 1) And alphaTile(char) <> 0 'Custom if a-z <> 1-26 and <> 0
30        ElseIf char = 32 Then
40            isCharacterCustom = (alphaTile(char) <> 0) 'Custom if spacebar different than 0
          
50        Else
60            isCharacterCustom = (alphaTile(char) <> 0) 'Custom if defined
70        End If
End Function


Private Sub Class_Initialize()
      'Automatic definition of tiles considering A = tile1...
          Dim i As Integer
          Dim tileid As Integer

10        ReDim alphaTile(255) As Integer

20        tileid = 1
30        For i = Asc("a") To Asc("z")
40            alphaTile(i) = tileid
50            tileid = tileid + 1
60        Next

          'Define Spacebar as tile 0
70        alphaTile(32) = 0

          '''''

80        Boundaries.Left = 0
90        Boundaries.Right = 1024
100       Boundaries.Top = 0
110       Boundaries.Bottom = 1024

120       Active = False
130       insert = True

End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        curtilex = (parent.hScr.value + X) \ parent.currenttilew
20        curtiley = (parent.vScr.value + Y) \ parent.currenttilew

30        If Button = vbRightButton Then
              'Stop writing on right-click
40            Call StopTyping

50        ElseIf Button = vbLeftButton Then
60            If Active Then
70                If curtilex >= Boundaries.Left And curtilex <= Boundaries.Right And _
                     curtiley >= Boundaries.Top And curtiley <= Boundaries.Bottom Then
                      'Clicked in text area, move cursor
80                    dragStart = CoordToStringPosition(curtilex, curtiley)
90                    frmGeneral.TileTextData.selstart = dragStart
100                   frmGeneral.TileTextData.sellength = 0

110                   Call parent.UpdatePreview(True, False)
                      'We might want to start selecting stuff
120                   dragging = True

130               Else
                      'Clicked outside text area ; stop typing
140                   Call StopTyping
150               End If

160           Else
                  'Tiletext not active; start typing at given coordinates
170               Call StartTyping(curtilex, curtiley)
180           End If
190       End If

End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
          Dim curtilex As Integer
          Dim curtiley As Integer

10        curtilex = (parent.hScr.value + X) \ parent.currenttilew
20        curtiley = (parent.vScr.value + Y) \ parent.currenttilew

30        If dragging Then
40            dragEnd = CoordToStringPosition(curtilex, curtiley)

50            If dragEnd >= dragStart Then
60                frmGeneral.TileTextData.selstart = dragStart
70                frmGeneral.TileTextData.sellength = dragEnd - dragStart
80            Else
90                frmGeneral.TileTextData.selstart = dragEnd
100               frmGeneral.TileTextData.sellength = dragStart - dragEnd
110           End If

120           If parent.magnifier.UsingPixels Then
130               Call parent.UpdateLevel
140           Else
150               Call parent.UpdatePreview(True, False)
160           End If
170       End If
End Sub

Sub MouseUp(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        curtilex = (parent.hScr.value + X) \ parent.currenttilew
20        curtiley = (parent.vScr.value + Y) \ parent.currenttilew

30        If dragging = True Then
              '
40        End If

50        dragging = False

End Sub


Private Sub DrawTextArea(startx As Integer, startY As Integer, endX As Integer, endY As Integer)
      'Swap first/current values if needed
          Dim x1 As Integer
          Dim x2 As Integer
          Dim y1 As Integer
          Dim y2 As Integer

10        If startx < endX Then
20            x1 = startx
30            x2 = endX
40        Else
50            x1 = endX
60            x2 = startx
70        End If

80        If startY < endY Then
90            y1 = startY
100           y2 = endY
110       Else
120           y1 = endY
130           y2 = startY
140       End If

          'update the drawing rectangle
150       If parent.magnifier.UsingPixels Then
160           parent.shptext.Left = (x1 * parent.currenttilew) - parent.hScr.value - 1
170           parent.shptext.Top = (y1 * parent.currenttilew) - parent.vScr.value - 1
180           parent.shptext.width = (x2 - x1 + 1) * parent.currenttilew + 2
190           parent.shptext.height = (y2 - y1 + 1) * parent.currenttilew + 2
200       Else
210           parent.shptext.Left = (x1 * parent.currenttilew) - parent.hScr.value
220           parent.shptext.Top = (y1 * parent.currenttilew) - parent.vScr.value
230           parent.shptext.width = (x2 - x1 + 1) * parent.currenttilew
240           parent.shptext.height = (y2 - y1 + 1) * parent.currenttilew
250       End If

End Sub

Sub KeyPress(Key As Integer)
10        If Active Then
20            Call SendMessageLong(frmGeneral.TileTextData.hWnd, WM_CHAR, Key, 0)
30            On Error Resume Next
40            parent.picpreview.setfocus

50            lastkey = Key

60            If parent.magnifier.UsingPixels Then
70                Call parent.UpdateLevel
80            Else
90                Call parent.UpdatePreview(True, False)
100           End If
110       End If
End Sub

Sub KeyDown(KeyCode As Integer, Shift As Integer)
10        If Active Then
20            Call SendMessageLong(frmGeneral.TileTextData.hWnd, WM_KEYDOWN, KeyCode, Shift)
30            On Error Resume Next
40            parent.picpreview.setfocus

50            lastkey = KeyCode

60            If parent.magnifier.UsingPixels Then
70                Call parent.UpdateLevel
                  'Call UpdateTextPreview(True, True)
80            Else
                  'Call parent.UpdatePreview(True, False)
90                Call parent.UpdatePreview(True, False)
100           End If

110       End If
End Sub

Sub KeyUp(KeyCode As Integer, Shift As Integer)
10        If Active Then
20            Call SendMessageLong(frmGeneral.TileTextData.hWnd, WM_KEYUP, KeyCode, Shift)
30            On Error Resume Next
40            parent.picpreview.setfocus
50        End If
End Sub

Sub StopTyping()
10        Active = False
20        parent.shptext.visible = False

30        Call UpdateTextPreview(False, True)

      '    Call parent.UpdateLevel
End Sub


Sub StartTyping(X As Integer, Y As Integer)

10        Active = True

20        Boundaries.Left = X
30        Boundaries.Right = X
40        Boundaries.Top = Y
50        Boundaries.Bottom = Y

          'update the drawing rectangle
60        Call DrawTextArea(Boundaries.Left, Boundaries.Top, Boundaries.Right, Boundaries.Bottom)

70        parent.shptext.visible = True

          'reset text
80        frmGeneral.TileTextData.Text = ""

90        Call DrawCursor(Boundaries.Left, Boundaries.Top)

End Sub





Private Sub DrawCursor(tileX As Integer, tileY As Integer)
10        If Active Then
20            parent.shpcursor.visible = True
30            parent.shpcursor.Left = (tileX * parent.currenttilew) - parent.hScr.value
40            parent.shpcursor.Top = (tileY * parent.currenttilew) - parent.vScr.value

50            If parent.magnifier.UsingPixels Then
60                parent.shpcursor.width = 1
70                parent.shpcursor.height = 3
80                parent.shpcursor.Top = parent.shpcursor.Top - 1
90            Else
100               parent.shpcursor.width = 2
110               parent.shpcursor.height = parent.currenttilew
120           End If
130       Else
140           parent.shpcursor.visible = False
150           parent.shpcursor.width = parent.currenttilew
160       End If

End Sub

Public Property Get isActive() As Boolean
10        isActive = Active
End Property

Public Property Let isActive(ByVal newActive As Boolean)
10        Active = newActive
End Property


Sub UpdateTextPreview(Optional onpreview As Boolean = True, Optional setfocus As Boolean = False)
          Dim typed As String
          Dim selstart As Integer
          Dim sellength As Integer

          Static stringlength As Integer

          'gather information from the textbox
10        typed = frmGeneral.TileTextData.Text
20        selstart = frmGeneral.TileTextData.selstart
30        sellength = frmGeneral.TileTextData.sellength
          
40        If stringlength > Len(typed) Or ((lastkey = 13 Or lastkey = 10) And selstart < stringlength) Then
50            stringlength = Len(typed)
60            lastkey = 0
70            Call parent.UpdateLevel
80        End If

90        stringlength = Len(typed)

100       Call DisplayChars(typed, selstart, sellength, onpreview, setfocus)

110       Call DrawTextArea(Boundaries.Left, Boundaries.Top, Boundaries.Right, Boundaries.Bottom)

120       If Not onpreview Or Boundaries.Right > 1024 Or Boundaries.Bottom > 1023 Then
130           If Boundaries.Right >= 1024 Then
140               Boundaries.Left = Boundaries.Left - (Boundaries.Right - 1024)
150               Boundaries.Right = 1024
160               If Boundaries.Left < 0 Then
                      'in case someone really crazy would input > 1024 characters on the same line
170                   Boundaries.Left = 0
180                   frmGeneral.TileTextData.Text = Mid(frmGeneral.TileTextData.Text, 1, 1023)
190                   Call StopTyping
200                   Exit Sub
210               End If
220           End If

230           If Boundaries.Bottom >= 1023 Then
240               Boundaries.Top = Boundaries.Top - (Boundaries.Bottom - 1023)
250               Boundaries.Bottom = 1023
260               If Boundaries.Top < 0 Then
                      'in case someone REAAAAAALLY crazy would input > 1024 LINES...
270                   Boundaries.Top = 0
280                   frmGeneral.TileTextData.Text = Mid(frmGeneral.TileTextData.Text, 1, 1023)
290                   Call StopTyping
300                   Exit Sub
310               End If
320           End If

330           Call parent.UpdateLevel(False)

340       End If

End Sub

Private Sub DisplayChars(str As String, selstart As Integer, sellength As Integer, onpreview As Boolean, setfocus As Boolean)
          Dim i As Integer

          Dim curtilex As Integer
          Dim curtiley As Integer

          Dim cursorX As Integer
          Dim cursorY As Integer

          Dim char As Integer

          Dim undoch As Changes

          Dim tilesDrawn As Boolean
          
          Dim lbx As Integer    'lowestval for x
          Dim hbx As Integer    'highestval for x
          Dim lby As Integer    'lowestval for y
          Dim hby As Integer    'highestval for y
10        lbx = parent.hScr.value \ parent.currenttilew
20        hbx = (parent.hScr.value + parent.picpreview.width) \ parent.currenttilew
30        lby = parent.vScr.value \ parent.currenttilew
40        hby = (parent.vScr.value + parent.picpreview.height) \ parent.currenttilew


50        If Not onpreview Then
60            parent.undoredo.ResetRedo
70            Set undoch = New Changes
80        End If
          
90        Boundaries.Right = Boundaries.Left
100       Boundaries.Bottom = Boundaries.Top

110       curtilex = Boundaries.Left
120       curtiley = Boundaries.Top

130       For i = 1 To Len(str)

140           If i - 1 = selstart Then
150               cursorX = curtilex
160               cursorY = curtiley
170           End If

180           char = CheckCase(Asc(Mid(str, i, 1)))
190           If char = 13 Then
                  'enter
200               curtilex = Boundaries.Left
210               curtiley = curtiley + 1

220               If curtiley > Boundaries.Bottom Then Boundaries.Bottom = curtiley

230           ElseIf char <> 10 Then    'ignore chr(10)

240               Call DrawTextTile(alphaTile(char), curtilex, curtiley, i, selstart, sellength, onpreview, undoch)
250               tilesDrawn = True
                  
260               If char >= 32 Then

270                   curtilex = curtilex + 1

280                   If curtilex > Boundaries.Right Then Boundaries.Right = curtilex

290               End If
300           End If
310       Next

320       If i - 1 = selstart Then
330           cursorX = curtilex
340           cursorY = curtiley
350           If curtiley > Boundaries.Bottom Then Boundaries.Bottom = curtiley
360           If curtilex > Boundaries.Right Then Boundaries.Right = curtilex
370       End If

380       Call DrawCursor(cursorX, cursorY)
          
390       If setfocus Then
400           If (cursorX <= lbx Or cursorX >= hbx Or cursorY <= lby Or cursorY >= hby) And _
                 cursorX >= 0 And cursorX <= 1023 And cursorY >= 0 And cursorY <= 1023 Then
410               Call parent.SetFocusAt(cursorX, cursorY, parent.picpreview.width \ 2, parent.picpreview.height \ 2, True)
420           End If
430       End If
          
440       On Error Resume Next
450       parent.picpreview.setfocus

460       If Not onpreview And tilesDrawn Then
470           Call parent.undoredo.AddToUndo(undoch, UNDO_TILETEXT)
480       End If

End Sub


'If a lowercase character is entered and that no tile is assigned to it,
'it will check if a tile is defined for the uppercase character, if yes, it returns
'that character ; same for upper to lower case
Private Function CheckCase(char As Integer) As Integer
10        CheckCase = char
20        If IsLcase(char) And alphaTile(char) = 0 Then
30            If alphaTile(char - Asc("a") + Asc("A")) <> 0 Then
40                CheckCase = char - Asc("a") + Asc("A")
50            End If
60        ElseIf IsUcase(char) And alphaTile(char) = 0 Then
70            If alphaTile(char + Asc("a") - Asc("A")) <> 0 Then
80                CheckCase = char + Asc("a") - Asc("A")
90            End If
100       End If
End Function

Private Sub DrawTextTile(tilenr As Integer, X As Integer, Y As Integer, strpos As Integer, selstart As Integer, sellength As Integer, onpreview As Boolean, Optional undoch As Changes, Optional drawblack As Boolean = True)

10        If onpreview Then
              Dim selected As Boolean
20            selected = sellength >= 1 And strpos - 1 >= selstart And strpos <= (selstart + sellength)
30            If parent.magnifier.UsingPixels Then
                  Dim color As Long
40                If selected Then
50                    color = vbWhite
60                Else
70                    color = TilePixelColor(tilenr)
80                End If
                  
90                Call SetPixel(parent.picpreview.hDC, (X - (parent.hScr.value \ parent.currenttilew)), (Y - (parent.vScr.value \ parent.currenttilew)), color)
100           Else
110                   If selected Then
120               frmGeneral.Label6.Caption = strpos & " selected"
130           End If
              
140               Call parent.TileRender.DrawTile(tilenr, selected, parent.picpreview.hDC, X * parent.currenttilew - parent.hScr.value, Y * parent.currenttilew - parent.vScr.value, True, drawblack)
150           End If
              

160       Else
170           If parent.sel.getIsInSelection(X, Y) Then
180               If TileIsSpecial(parent.sel.getSelTile(X, Y)) Then
190                   Call SearchAndDestroySelObject(parent.sel, X, Y, undoch)
200               End If
210               Call parent.sel.setSelTile(X, Y, tilenr, undoch, True)
220           Else
230               If TileIsSpecial(parent.getTile(X, Y)) Then
240                   Call SearchAndDestroyObject(parent, X, Y, undoch, False)
250               End If
260               Call parent.setTile(X, Y, tilenr, undoch, True)
270               Call parent.UpdateLevelTile(X, Y, False, True)
280           End If
290       End If

End Sub


'Returns corresponding position in the string of a x/y coordinate
Private Function CoordToStringPosition(X As Integer, Y As Integer) As Integer
          Dim lines() As String
          Dim i As Integer
          Dim strpos As Integer
10        strpos = 0

20        If X < Boundaries.Left Then X = Boundaries.Left
30        If X > Boundaries.Right Then X = Boundaries.Right
40        If Y < Boundaries.Top Then Y = Boundaries.Top
50        If Y > Boundaries.Bottom Then Y = Boundaries.Bottom

60        lines = Split(frmGeneral.TileTextData.Text, Chr(13) & Chr(10))

70        If Y > Boundaries.Top Then
80            For i = 0 To Y - Boundaries.Top - 1
90                strpos = strpos + Len(lines(i)) + 2
100           Next
110       End If

120       If (X - Boundaries.Left) > Len(lines(Y - Boundaries.Top)) Then
130           strpos = strpos + Len(lines(Y - Boundaries.Top))
140       Else
150           strpos = strpos + X - Boundaries.Left
160       End If

170       CoordToStringPosition = strpos

End Function

Private Sub Class_Terminate()
10        Set parent = Nothing
End Sub
