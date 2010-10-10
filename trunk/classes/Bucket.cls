VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain
Dim c_isfilling As Boolean
Dim undoch As Changes


Public Property Get isFilling() As Boolean
    isFilling = c_isfilling
End Property

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    On Error GoTo MouseDown_Error

    curtilex = (parent.hScr.value + X) \ parent.currenttilew
    curtiley = (parent.vScr.value + Y) \ parent.currenttilew
    
    'Boundary check
    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
        Exit Sub
    End If

    Dim curSelection As TilesetSelection
    curSelection = parent.tileset.selection(Button)

    'Unsupported bucket features
    If curSelection.isSpecialObject Or _
        curSelection.selectionType = TS_Block Or _
        curSelection.selectionType = TS_LVZ Then

        Call MessageBox("Bucket fill is not supported for this type of object.", vbOKOnly + vbInformation)
        Exit Sub
    End If

    Dim fillinselectiononly As Boolean
    fillinselectiononly = (parent.sel.hasAlreadySelectedParts And parent.sel.getIsInSelection(curtilex, curtiley))



    Dim sourcetilenr As Integer
    If parent.sel.getIsInSelection(curtilex, curtiley) Then
        sourcetilenr = parent.sel.getSelTile(curtilex, curtiley)
    Else
        sourcetilenr = parent.getTile(curtilex, curtiley)
    End If

    If sourcetilenr < 0 Then
        sourcetilenr = sourcetilenr \ -100
    End If
    
    'No need to do anything if the source tile is the same as target
    If curSelection.selectionType <> TS_Walltiles And sourcetilenr = curSelection.tilenr Then
        Exit Sub
    End If
    
    'Initialize changes stack
    parent.undoredo.ResetRedo
    Set undoch = New Changes
    
    If curSelection.selectionType = TS_Walltiles Then
        Call parent.walltiles.ClearIsWallTile
        Call Fill(sourcetilenr, curSelection, curtilex, curtiley, undoch, fillinselectiononly, frmGeneral.chkFillInScreen.value = vbChecked, True)
    End If
    Call Fill(sourcetilenr, curSelection, curtilex, curtiley, undoch, fillinselectiononly, frmGeneral.chkFillInScreen.value = vbChecked, False)

    Call parent.undoredo.AddToUndo(undoch, UNDO_BUCKETFILL)

    On Error GoTo 0
    Exit Sub

MouseDown_Error:
    HandleError Err, "Bucket.MouseDown"
End Sub

Sub MouseMove()
'no way we're using bucket at move, it will crash like hell
End Sub

Sub MouseUp()
'nothing to do on up
End Sub

Friend Sub Fill(tilenr As Integer, ByRef curSelection As TilesetSelection, ByVal X As Integer, ByVal Y As Integer, undoch As Changes, fillinselection As Boolean, fillscreenonly As Boolean, Optional tagwalltiles As Boolean = False)
    On Error GoTo Fill_Error

    Dim tick As Long
    tick = GetTickCount
    
    frmGeneral.IsBusy("Bucket.Fill") = True
    
    c_isfilling = True

    Dim done(1023, 1023) As Boolean
    Dim px() As Long
    Dim py() As Long
    Dim Pstart As Long
    Dim Pstop As Long
    Dim Pcount As Long
    
    Dim Colour As Integer, fillnr As Integer
    
    fillnr = curSelection.tilenr
    
    Dim usewalltile As Boolean
    usewalltile = (curSelection.selectionType = TS_Walltiles)
      
    Dim lX As Integer
    Dim lY As Integer

    Dim Sloop As Long
    ReDim px(0)
    ReDim py(0)

    lX = X
    lY = Y

    Pstart = 0
    Pstop = 0
    px(0) = lX
    py(0) = lY
    Dim count As Long

    Dim lbsx As Integer
    Dim lbsy As Integer
    Dim hbsx As Integer
    Dim hbsy As Integer
    lbsx = parent.hScr.value \ parent.currenttilew
    hbsx = (parent.hScr.value + parent.picPreview.width) \ parent.currenttilew
    lbsy = parent.vScr.value \ parent.currenttilew
    hbsy = (parent.vScr.value + parent.picPreview.Height) \ parent.currenttilew

    Call AddPoints(px, py, Pstart, Pstop, lX, lY, Sloop)

    While Pstart <= Pstop And count < MAX_BUCKETFILL

        lX = px(Pstart)
        lY = py(Pstart)
        If Not (lX < 0 Or lY < 0 Or lX > 1023 Or lY > 1023) Then
            If fillinselection Then
                If Not parent.sel.getIsInSelection(lX, lY) Then
                    GoTo skiploop
                End If
            ElseIf Not fillinselection And fillscreenonly Then
                If (lX < lbsx Or lY < lbsy Or lX > hbsx Or lY > hbsy) Then
                    GoTo skiploop
                End If
            End If

            If parent.sel.getIsInSelection(lX, lY) Then
                Colour = parent.sel.getSelTile(lX, lY)
            Else
                Colour = parent.getTile(lX, lY)
            End If

            If Colour < 0 Then
                Colour = Colour \ -100
            End If

            If Not (done(lX, lY) Or Colour <> tilenr) Then
                Pcount = Pcount + 1
                done(lX, lY) = True

                count = count + 1

                If tagwalltiles Then
                    Call parent.walltiles.setIsWallTile(lX, lY)
                Else

                    If fillinselection Then
                        If usewalltile Then
                            fillnr = parent.walltiles.DrawSelWallTile(lX, lY, undoch, False, False)
                        End If
                        Call parent.sel.setSelTile(lX, lY, fillnr, undoch)
                    Else
                        If usewalltile Then
                            fillnr = parent.walltiles.DrawWallTile(lX, lY, False, undoch, False, True)
                        End If
                        Call parent.setTile(lX, lY, fillnr, undoch, True)
                    End If
                End If
                'it will prolly go faster if we directly set the pixel to the
                'pixel map, rather than calling updateleveltile,
                'becuase updateleveltile also calculates if it falls
                'in preview, which in unnecessary

                'Call SetPixel(parent.pic1024.hdc, Lx, Ly, TilePixelColor(parent.getTile(x, y)))
                'Call parent.Set1024Pixel(Lx, Ly, TilePixelColor(parent.getTile(X, Y)))
                If Not fillinselection Then
                    Call parent.UpdateLevelTile(lX, lY, False)
                End If
                '                piccurrent.PSet (Lx, Ly), vbblue
                '  DoEvents
                Call AddPoints(px, py, Pstart, Pstop, lX, lY, Sloop)
            End If
        End If
skiploop:
        Pstart = Pstart + 1

        If count >= MAX_BUCKETFILL Then
          
            parent.AddDebug "Bucket.Fill, " & count & " tiles filled in " & GetTickCount - tick & " ms"
            
            Call parent.UpdateLevel
            If MessageBox(MAX_BUCKETFILL & " tiles were filled. You might be trying to fill the entire map, which would cause DCME to freeze for several minutes. Select 'OK' to abort the operation, or 'Cancel' to continue to fill the area.", vbOKCancel + vbExclamation) = vbCancel Then
                count = 0
                
                tick = GetTickCount
            End If
            
        End If
        
    Wend
    'but because we handled it directly, we need to update the level
    'too,not only the preview

    

    If count >= MAX_BUCKETFILL Then
        MessageBox "Fill area too large, operation aborted." & vbCrLf & "(" & count & " tiles filled)", vbCritical + vbOKOnly, "Fill aborted"
    Else
        'We already added this info to the log if count >= MAX_BUCKETFILL
        parent.AddDebug "Bucket.Fill, " & count & " tiles filled in " & GetTickCount - tick & " ms"
    End If

    If Not tagwalltiles Then Call parent.UpdateLevel
    c_isfilling = False

    frmGeneral.IsBusy("Bucket.Fill") = False

    On Error GoTo 0
    Exit Sub

Fill_Error:

    frmGeneral.IsBusy("Bucket.Fill") = False
    
    HandleError Err, "Bucket.Fill"
End Sub

Private Sub AddPoints(ByRef px() As Long, ByRef py() As Long, ByRef Pstart As Long, ByRef Pstop As Long, ByRef lX As Integer, ByRef lY As Integer, ByRef Sloop As Long)

    If Pstart > BUCKET_STACK Then
        For Sloop = Pstart To Pstop
            px(Sloop - BUCKET_STACK) = px(Sloop)
            py(Sloop - BUCKET_STACK) = py(Sloop)
        Next Sloop
        ReDim Preserve px(Pstop - BUCKET_STACK)
        ReDim Preserve py(Pstop - BUCKET_STACK)
        Pstart = Pstart - BUCKET_STACK
        Pstop = Pstop - BUCKET_STACK
    End If
    
    If frmGeneral.chkFillDiagonal.value = vbChecked Then
        ReDim Preserve px(Pstop + 8)
        ReDim Preserve py(Pstop + 8)
        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY

        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY + 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY + 1

        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY + 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY

        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY - 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY - 1

        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY - 1



    Else
        ReDim Preserve px(Pstop + 4)
        ReDim Preserve py(Pstop + 4)
        Pstop = Pstop + 1
        px(Pstop) = lX + 1
        py(Pstop) = lY
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY + 1
        Pstop = Pstop + 1
        px(Pstop) = lX - 1
        py(Pstop) = lY
        Pstop = Pstop + 1
        px(Pstop) = lX
        py(Pstop) = lY - 1
    End If

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
