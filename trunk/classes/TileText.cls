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
    Set parent = Main
End Sub

Function GetalphaTiles() As Integer()
    GetalphaTiles = alphaTile
End Function

Function GetalphaTile(char As Integer) As Integer
    GetalphaTile = alphaTile(char)
End Function

Sub SetalphaTiles(newalphatile() As Integer)
    alphaTile() = newalphatile()
End Sub

Sub SetalphaTile(char As Integer, tile As Integer)
    alphaTile(char) = tile
End Sub

Function isCharacterCustom(char As Integer) As Boolean
    'Returns true if the tile defined for the specified character is different from default
    If char >= Asc("a") And char <= Asc("z") Then
        isCharacterCustom = (alphaTile(char) <> char - Asc("a") + 1) And alphaTile(char) <> 0 'Custom if a-z <> 1-26 and <> 0
    ElseIf char = 32 Then
        isCharacterCustom = (alphaTile(char) <> 0) 'Custom if spacebar different than 0
    
    Else
        isCharacterCustom = (alphaTile(char) <> 0) 'Custom if defined
    End If
End Function


Private Sub Class_Initialize()
'Automatic definition of tiles considering A = tile1...
    Dim i As Integer
    Dim tileid As Integer

    ReDim alphaTile(255) As Integer

    tileid = 1
    For i = Asc("a") To Asc("z")
        alphaTile(i) = tileid
        tileid = tileid + 1
    Next

    'Define Spacebar as tile 0
    alphaTile(32) = 0

    '''''

    Boundaries.Left = 0
    Boundaries.Right = 1024
    Boundaries.Top = 0
    Boundaries.Bottom = 1024

    Active = False
    insert = True

End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
    Dim curtilex As Integer
    Dim curtiley As Integer
    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If Button = vbRightButton Then
        'Stop writing on right-click
        Call StopTyping

    ElseIf Button = vbLeftButton Then
        If Active Then
            If curtilex >= Boundaries.Left And curtilex <= Boundaries.Right And _
               curtiley >= Boundaries.Top And curtiley <= Boundaries.Bottom Then
                'Clicked in text area, move cursor
                dragStart = CoordToStringPosition(curtilex, curtiley)
                frmGeneral.TileTextData.selstart = dragStart
                frmGeneral.TileTextData.sellength = 0

                Call parent.UpdatePreview(True, False)
                'We might want to start selecting stuff
                dragging = True

            Else
                'Clicked outside text area ; stop typing
                Call StopTyping
            End If

        Else
            'Tiletext not active; start typing at given coordinates
            Call StartTyping(curtilex, curtiley)
        End If
    End If

End Sub

Sub MouseMove(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
    Dim curtilex As Integer
    Dim curtiley As Integer

    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If dragging Then
        dragEnd = CoordToStringPosition(curtilex, curtiley)

        If dragEnd >= dragStart Then
            frmGeneral.TileTextData.selstart = dragStart
            frmGeneral.TileTextData.sellength = dragEnd - dragStart
        Else
            frmGeneral.TileTextData.selstart = dragEnd
            frmGeneral.TileTextData.sellength = dragStart - dragEnd
        End If

        If parent.magnifier.UsingPixels Then
            Call parent.UpdateLevel
        Else
            Call parent.UpdatePreview(True, False)
        End If
    End If
End Sub

Sub MouseUp(Button As Integer, X As Single, Y As Single, Optional Shift As Integer)
    Dim curtilex As Integer
    Dim curtiley As Integer
    curtilex = (parent.Hscr.value + X) \ parent.currenttilew
    curtiley = (parent.Vscr.value + Y) \ parent.currenttilew

    If dragging = True Then
        '
    End If

    dragging = False

End Sub


Private Sub DrawTextArea(startx As Integer, startY As Integer, endX As Integer, endY As Integer)
'Swap first/current values if needed
    Dim x1 As Integer
    Dim x2 As Integer
    Dim y1 As Integer
    Dim y2 As Integer

    If startx < endX Then
        x1 = startx
        x2 = endX
    Else
        x1 = endX
        x2 = startx
    End If

    If startY < endY Then
        y1 = startY
        y2 = endY
    Else
        y1 = endY
        y2 = startY
    End If

    'update the drawing rectangle
    If parent.magnifier.UsingPixels Then
        parent.shptext.Left = (x1 * parent.currenttilew) - parent.Hscr.value - 1
        parent.shptext.Top = (y1 * parent.currenttilew) - parent.Vscr.value - 1
        parent.shptext.width = (x2 - x1 + 1) * parent.currenttilew + 2
        parent.shptext.height = (y2 - y1 + 1) * parent.currenttilew + 2
    Else
        parent.shptext.Left = (x1 * parent.currenttilew) - parent.Hscr.value
        parent.shptext.Top = (y1 * parent.currenttilew) - parent.Vscr.value
        parent.shptext.width = (x2 - x1 + 1) * parent.currenttilew
        parent.shptext.height = (y2 - y1 + 1) * parent.currenttilew
    End If

End Sub

Sub KeyPress(Key As Integer)
    If Active Then
        Call SendMessageLong(frmGeneral.TileTextData.hWnd, WM_CHAR, Key, 0)
        On Error Resume Next
        parent.picPreview.setfocus

        lastkey = Key

        If parent.magnifier.UsingPixels Then
            Call parent.UpdateLevel
        Else
            Call parent.UpdatePreview(True, False)
        End If
    End If
End Sub

Sub KeyDown(KeyCode As Integer, Shift As Integer)
    If Active Then
        Call SendMessageLong(frmGeneral.TileTextData.hWnd, WM_KEYDOWN, KeyCode, Shift)
        On Error Resume Next
        parent.picPreview.setfocus

        lastkey = KeyCode

        If parent.magnifier.UsingPixels Then
            Call parent.UpdateLevel
            'Call UpdateTextPreview(True, True)
        Else
            'Call parent.UpdatePreview(True, False)
            Call parent.UpdatePreview(True, False)
        End If

    End If
End Sub

Sub KeyUp(KeyCode As Integer, Shift As Integer)
    If Active Then
        Call SendMessageLong(frmGeneral.TileTextData.hWnd, WM_KEYUP, KeyCode, Shift)
        On Error Resume Next
        parent.picPreview.setfocus
    End If
End Sub

Sub StopTyping()
    Active = False
    parent.shptext.visible = False

    Call UpdateTextPreview(False, True)

'    Call parent.UpdateLevel
End Sub


Sub StartTyping(X As Integer, Y As Integer)

    Active = True

    Boundaries.Left = X
    Boundaries.Right = X
    Boundaries.Top = Y
    Boundaries.Bottom = Y

    'update the drawing rectangle
    Call DrawTextArea(Boundaries.Left, Boundaries.Top, Boundaries.Right, Boundaries.Bottom)

    parent.shptext.visible = True

    'reset text
    frmGeneral.TileTextData.Text = ""

    Call DrawCursor(Boundaries.Left, Boundaries.Top)

End Sub





Private Sub DrawCursor(tileX As Integer, tileY As Integer)
    If Active Then
        parent.shpcursor.visible = True
        parent.shpcursor.Left = (tileX * parent.currenttilew) - parent.Hscr.value
        parent.shpcursor.Top = (tileY * parent.currenttilew) - parent.Vscr.value

        If parent.magnifier.UsingPixels Then
            parent.shpcursor.width = 1
            parent.shpcursor.height = 3
            parent.shpcursor.Top = parent.shpcursor.Top - 1
        Else
            parent.shpcursor.width = 2
            parent.shpcursor.height = parent.currenttilew
        End If
    Else
        parent.shpcursor.visible = False
        parent.shpcursor.width = parent.currenttilew
    End If

End Sub

Public Property Get isActive() As Boolean
    isActive = Active
End Property

Public Property Let isActive(ByVal newActive As Boolean)
    Active = newActive
End Property


Sub UpdateTextPreview(Optional onpreview As Boolean = True, Optional setfocus As Boolean = False)
    Dim typed As String
    Dim selstart As Integer
    Dim sellength As Integer

    Static stringlength As Integer

    'gather information from the textbox
    typed = frmGeneral.TileTextData.Text
    selstart = frmGeneral.TileTextData.selstart
    sellength = frmGeneral.TileTextData.sellength
    
    If stringlength > Len(typed) Or ((lastkey = 13 Or lastkey = 10) And selstart < stringlength) Then
        stringlength = Len(typed)
        lastkey = 0
        Call parent.UpdateLevel
    End If

    stringlength = Len(typed)

    Call DisplayChars(typed, selstart, sellength, onpreview, setfocus)

    Call DrawTextArea(Boundaries.Left, Boundaries.Top, Boundaries.Right, Boundaries.Bottom)

    If Not onpreview Or Boundaries.Right > 1024 Or Boundaries.Bottom > 1023 Then
        If Boundaries.Right >= 1024 Then
            Boundaries.Left = Boundaries.Left - (Boundaries.Right - 1024)
            Boundaries.Right = 1024
            If Boundaries.Left < 0 Then
                'in case someone really crazy would input > 1024 characters on the same line
                Boundaries.Left = 0
                frmGeneral.TileTextData.Text = Mid(frmGeneral.TileTextData.Text, 1, 1023)
                Call StopTyping
                Exit Sub
            End If
        End If

        If Boundaries.Bottom >= 1023 Then
            Boundaries.Top = Boundaries.Top - (Boundaries.Bottom - 1023)
            Boundaries.Bottom = 1023
            If Boundaries.Top < 0 Then
                'in case someone REAAAAAALLY crazy would input > 1024 LINES...
                Boundaries.Top = 0
                frmGeneral.TileTextData.Text = Mid(frmGeneral.TileTextData.Text, 1, 1023)
                Call StopTyping
                Exit Sub
            End If
        End If

        Call parent.UpdateLevel(False)

    End If

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
    lbx = parent.Hscr.value \ parent.currenttilew
    hbx = (parent.Hscr.value + parent.picPreview.width) \ parent.currenttilew
    lby = parent.Vscr.value \ parent.currenttilew
    hby = (parent.Vscr.value + parent.picPreview.height) \ parent.currenttilew


    If Not onpreview Then
        parent.undoredo.ResetRedo
        Set undoch = New Changes
    End If
    
    Boundaries.Right = Boundaries.Left
    Boundaries.Bottom = Boundaries.Top

    curtilex = Boundaries.Left
    curtiley = Boundaries.Top

    For i = 1 To Len(str)

        If i - 1 = selstart Then
            cursorX = curtilex
            cursorY = curtiley
        End If

        char = CheckCase(Asc(Mid(str, i, 1)))
        If char = 13 Then
            'enter
            curtilex = Boundaries.Left
            curtiley = curtiley + 1

            If curtiley > Boundaries.Bottom Then Boundaries.Bottom = curtiley

        ElseIf char <> 10 Then    'ignore chr(10)

            Call DrawTextTile(alphaTile(char), curtilex, curtiley, i, selstart, sellength, onpreview, undoch)
            tilesDrawn = True
            
            If char >= 32 Then

                curtilex = curtilex + 1

                If curtilex > Boundaries.Right Then Boundaries.Right = curtilex

            End If
        End If
    Next

    If i - 1 = selstart Then
        cursorX = curtilex
        cursorY = curtiley
        If curtiley > Boundaries.Bottom Then Boundaries.Bottom = curtiley
        If curtilex > Boundaries.Right Then Boundaries.Right = curtilex
    End If

    Call DrawCursor(cursorX, cursorY)
    
    If setfocus Then
        If (cursorX <= lbx Or cursorX >= hbx Or cursorY <= lby Or cursorY >= hby) And _
           cursorX >= 0 And cursorX <= 1023 And cursorY >= 0 And cursorY <= 1023 Then
            Call parent.SetFocusAt(cursorX, cursorY, parent.picPreview.width \ 2, parent.picPreview.height \ 2, True)
        End If
    End If
    
    On Error Resume Next
    parent.picPreview.setfocus

    If Not onpreview And tilesDrawn Then
        Call parent.undoredo.AddToUndo(undoch, UNDO_TILETEXT)
    End If

End Sub


'If a lowercase character is entered and that no tile is assigned to it,
'it will check if a tile is defined for the uppercase character, if yes, it returns
'that character ; same for upper to lower case
Private Function CheckCase(char As Integer) As Integer
    CheckCase = char
    If IsLcase(char) And alphaTile(char) = 0 Then
        If alphaTile(char - Asc("a") + Asc("A")) <> 0 Then
            CheckCase = char - Asc("a") + Asc("A")
        End If
    ElseIf IsUcase(char) And alphaTile(char) = 0 Then
        If alphaTile(char + Asc("a") - Asc("A")) <> 0 Then
            CheckCase = char + Asc("a") - Asc("A")
        End If
    End If
End Function

Private Sub DrawTextTile(tilenr As Integer, X As Integer, Y As Integer, strpos As Integer, selstart As Integer, sellength As Integer, onpreview As Boolean, Optional undoch As Changes, Optional drawblack As Boolean = True)

    If onpreview Then
        Dim selected As Boolean
        selected = sellength >= 1 And strpos - 1 >= selstart And strpos <= (selstart + sellength)
        If parent.magnifier.UsingPixels Then
            Dim color As Long
            If selected Then
                color = vbWhite
            Else
                color = TilePixelColor(tilenr)
            End If
            
            Call SetPixel(parent.picPreview.hDC, (X - (parent.Hscr.value \ parent.currenttilew)), (Y - (parent.Vscr.value \ parent.currenttilew)), color)
        Else
                If selected Then
            frmGeneral.Label6.Caption = strpos & " selected"
        End If
        
            Call parent.TileRender.DrawTile(tilenr, selected, parent.picPreview.hDC, X * parent.currenttilew - parent.Hscr.value, Y * parent.currenttilew - parent.Vscr.value, True, drawblack)
        End If
        

    Else
        If parent.sel.getIsInSelection(X, Y) Then
            If TileIsSpecial(parent.sel.getSelTile(X, Y)) Then
                Call SearchAndDestroySelObject(parent.sel, X, Y, undoch)
            End If
            Call parent.sel.setSelTile(X, Y, tilenr, undoch, True)
        Else
            If TileIsSpecial(parent.getTile(X, Y)) Then
                Call SearchAndDestroyObject(parent, X, Y, undoch, False)
            End If
            Call parent.setTile(X, Y, tilenr, undoch, True)
            Call parent.UpdateLevelTile(X, Y, False, True)
        End If
    End If

End Sub


'Returns corresponding position in the string of a x/y coordinate
Private Function CoordToStringPosition(X As Integer, Y As Integer) As Integer
    Dim lines() As String
    Dim i As Integer
    Dim strpos As Integer
    strpos = 0

    If X < Boundaries.Left Then X = Boundaries.Left
    If X > Boundaries.Right Then X = Boundaries.Right
    If Y < Boundaries.Top Then Y = Boundaries.Top
    If Y > Boundaries.Bottom Then Y = Boundaries.Bottom

    lines = Split(frmGeneral.TileTextData.Text, Chr(13) & Chr(10))

    If Y > Boundaries.Top Then
        For i = 0 To Y - Boundaries.Top - 1
            strpos = strpos + Len(lines(i)) + 2
        Next
    End If

    If (X - Boundaries.Left) > Len(lines(Y - Boundaries.Top)) Then
        strpos = strpos + Len(lines(Y - Boundaries.Top))
    Else
        strpos = strpos + X - Boundaries.Left
    End If

    CoordToStringPosition = strpos

End Function

Private Sub Class_Terminate()
    Set parent = Nothing
End Sub
