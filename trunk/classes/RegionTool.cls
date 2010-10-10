VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegionTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim parent As frmMain
Dim UsingWand As Boolean
Dim firstX As Integer
Dim firstY As Integer
Dim drawing As Boolean

Dim undoch As Changes

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Sub MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, isWand As Boolean)
    Dim curtilex As Integer
    Dim curtiley As Integer

    curtilex = (parent.hScr.value + x) \ parent.currenttilew
    curtiley = (parent.vScr.value + y) \ parent.currenttilew

    If frmGeneral.llRegionList.ListIndex < 0 Then
        Exit Sub
    End If

    UsingWand = isWand

    If UsingWand Then
        'MAGIC WAND
        If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
            Exit Sub
        End If
        Set undoch = New Changes
        parent.undoredo.ResetRedo
        Call RegionsWand(Not IsControl(Shift) And Button = vbLeftButton, frmGeneral.llRegionList.ListIndex, IIf(parent.sel.getIsInSelection(curtilex, curtiley), parent.sel.getSelTile(curtilex, curtiley), parent.getTile(curtilex, curtiley)), curtilex, curtiley, undoch, True)
        
        Call parent.undoredo.AddToUndo(undoch, UNDO_REGION_ADD)
    Else
        'SELECTION

        If Button = vbLeftButton Or Button = vbRightButton Then

            If curtilex > 1023 Then curtilex = 1023
            If curtilex < 0 Then curtilex = 0
            If curtiley > 1023 Then curtiley = 1023
            If curtiley < 0 Then curtiley = 0

            firstX = curtilex
            firstY = curtiley

            parent.shpdraw.Left = firstX * parent.currenttilew - parent.hScr.value
            parent.shpdraw.Top = firstY * parent.currenttilew - parent.vScr.value
            parent.shpdraw.width = parent.currenttilew
            parent.shpdraw.height = parent.currenttilew

            parent.shpdraw.visible = True
            drawing = True

        End If

    End If


End Sub

Sub MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer

    curtilex = (parent.hScr.value + x) \ parent.currenttilew
    curtiley = (parent.vScr.value + y) \ parent.currenttilew

    If frmGeneral.llRegionList.ListIndex < 0 Then
        Exit Sub
    End If


    If UsingWand Then
        'MAGIC WAND
        If Button Then
            If curtilex > 1023 Then curtilex = 1023
            If curtilex < 0 Then curtilex = 0
            If curtiley > 1023 Then curtiley = 1023
            If curtiley < 0 Then curtiley = 0

            'do nothing

        End If
    Else
        'SELECTION
        If curtilex > 1023 Then curtilex = 1023
        If curtilex < 0 Then curtilex = 0
        If curtiley > 1023 Then curtiley = 1023
        If curtiley < 0 Then curtiley = 0


        If drawing Then
            If firstX <= curtilex Then
                parent.shpdraw.Left = firstX * parent.currenttilew - parent.hScr.value
                parent.shpdraw.width = (curtilex - firstX + 1) * parent.currenttilew
            Else
                parent.shpdraw.Left = curtilex * parent.currenttilew - parent.hScr.value
                parent.shpdraw.width = (firstX - curtilex + 1) * parent.currenttilew
            End If

            If firstY <= curtiley Then
                parent.shpdraw.Top = firstY * parent.currenttilew - parent.vScr.value
                parent.shpdraw.height = (curtiley - firstY + 1) * parent.currenttilew
            Else
                parent.shpdraw.Top = curtiley * parent.currenttilew - parent.vScr.value
                parent.shpdraw.height = (firstY - curtiley + 1) * parent.currenttilew
            End If

        End If
    End If

End Sub

Sub MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer

    curtilex = (parent.hScr.value + x) \ parent.currenttilew
    curtiley = (parent.vScr.value + y) \ parent.currenttilew

    If frmGeneral.llRegionList.ListIndex < 0 Then
        Exit Sub
    End If

    If UsingWand Then
        'MAGIC WAND

        'do nothing

    Else
        'SELECTION
        If curtilex > 1023 Then curtilex = 1023
        If curtilex < 0 Then curtilex = 0
        If curtiley > 1023 Then curtiley = 1023
        If curtiley < 0 Then curtiley = 0
        If drawing Then

            Set undoch = New Changes
            parent.undoredo.ResetRedo
            
            If IsControl(Shift) Or Button = vbRightButton Then
                Call parent.Regions.RemoveRectangleFromRegion(frmGeneral.llRegionList.ListIndex, firstX, firstY, curtilex, curtiley, undoch)
                Call parent.undoredo.AddToUndo(undoch, UNDO_REGION_REMOVE)
            Else
                Call parent.Regions.AddRectangleToRegion(frmGeneral.llRegionList.ListIndex, firstX, firstY, curtilex, curtiley, undoch)
                Call parent.undoredo.AddToUndo(undoch, UNDO_REGION_ADD)
            End If
            Call parent.RedrawRegions(True)

            parent.shpdraw.visible = False
            drawing = False
        End If
    End If


End Sub



Sub RegionsWand(add As Boolean, regionID As Integer, tilenr As Integer, ByVal x As Integer, ByVal y As Integer, undoch As Changes, Optional fillscreenonly = True, Optional ByVal stack As Long = 50000)

    On Error GoTo Regions_Error

    frmGeneral.IsBusy("RegionTool.RegionsWand") = True

    Dim done(1023, 1023) As Boolean
    Dim px() As Long
    Dim py() As Long
    Dim Pstart As Long
    Dim Pstop As Long
    Dim Pcount As Long
    Dim Colour As Integer

    Dim Lx As Integer
    Dim Ly As Integer

    Dim Sloop As Long
    ReDim px(0)
    ReDim py(0)

    Lx = x
    Ly = y

    If tilenr < 0 Then
        tilenr = tilenr \ -100
    End If

    Pstart = 0
    Pstop = 0
    px(0) = Lx
    py(0) = Ly
    Dim count As Long

    Dim lbsx As Integer
    Dim lbsy As Integer
    Dim hbsx As Integer
    Dim hbsy As Integer
    lbsx = parent.hScr.value \ parent.currenttilew
    hbsx = (parent.hScr.value + parent.picPreview.width) \ parent.currenttilew
    lbsy = parent.vScr.value \ parent.currenttilew
    hbsy = (parent.vScr.value + parent.picPreview.height) \ parent.currenttilew

    Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)

    While Pstart <= Pstop

        Lx = px(Pstart)
        Ly = py(Pstart)
        If Not (Lx < 0 Or Ly < 0 Or Lx > 1023 Or Ly > 1023) Then
            If fillscreenonly Then
                If (Lx < lbsx Or Ly < lbsy Or Lx > hbsx Or Ly > hbsy) Then
                    GoTo skiploop
                End If
            End If

            If parent.sel.getIsInSelection(Lx, Ly) Then
                Colour = parent.sel.getSelTile(Lx, Ly)
            Else
                Colour = parent.getTile(Lx, Ly)
            End If

            If Colour < 0 Then
                Colour = Colour \ -100
            End If

            If Not (done(Lx, Ly) = True Or Colour <> tilenr) Then
                Pcount = Pcount + 1
                done(Lx, Ly) = True

                count = count + 1

                If add Then
                    Call parent.Regions.AddToRegion(regionID, Lx, Ly, undoch)
                Else
                    Call parent.Regions.RemoveFromRegion(regionID, Lx, Ly, undoch)
                End If

                Call AddPoints(px, py, Pstart, Pstop, Lx, Ly, Sloop, stack)
            End If
        End If
skiploop:
        Pstart = Pstart + 1

    Wend

    Call parent.Regions.RedrawAllRegions
    
    Call parent.UpdateLevel

    frmGeneral.IsBusy("RegionTool.RegionsWand") = False

    On Error GoTo 0
    Exit Sub

Regions_Error:

    frmGeneral.IsBusy("RegionTool.RegionsWand") = False
    
    Call HandleError(Err, "RegionTool.RegionsWand", True, False)
End Sub

Private Sub AddPoints(ByRef px() As Long, ByRef py() As Long, ByRef Pstart As Long, ByRef Pstop As Long, ByRef Lx As Integer, ByRef Ly As Integer, ByRef Sloop As Long, ByRef stack As Long)

    If Pstart > stack Then
        For Sloop = Pstart To Pstop
            px(Sloop - stack) = px(Sloop)
            py(Sloop - stack) = py(Sloop)
        Next Sloop
        ReDim Preserve px(Pstop - stack)
        ReDim Preserve py(Pstop - stack)
        Pstart = Pstart - stack
        Pstop = Pstop - stack
    End If
    If False Then    'frmGeneral.chkMagicWandDiagonal.value = vbChecked
        ReDim Preserve px(Pstop + 8)
        ReDim Preserve py(Pstop + 8)
        Pstop = Pstop + 1
        px(Pstop) = Lx + 1
        py(Pstop) = Ly

        Pstop = Pstop + 1
        px(Pstop) = Lx + 1
        py(Pstop) = Ly + 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = Lx
        py(Pstop) = Ly + 1

        Pstop = Pstop + 1
        px(Pstop) = Lx - 1
        py(Pstop) = Ly + 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = Lx - 1
        py(Pstop) = Ly

        Pstop = Pstop + 1
        px(Pstop) = Lx - 1
        py(Pstop) = Ly - 1
        '------
        Pstop = Pstop + 1
        px(Pstop) = Lx
        py(Pstop) = Ly - 1

        Pstop = Pstop + 1
        px(Pstop) = Lx + 1
        py(Pstop) = Ly - 1



    Else
        ReDim Preserve px(Pstop + 4)
        ReDim Preserve py(Pstop + 4)
        Pstop = Pstop + 1
        px(Pstop) = Lx + 1
        py(Pstop) = Ly
        Pstop = Pstop + 1
        px(Pstop) = Lx
        py(Pstop) = Ly + 1
        Pstop = Pstop + 1
        px(Pstop) = Lx - 1
        py(Pstop) = Ly
        Pstop = Pstop + 1
        px(Pstop) = Lx
        py(Pstop) = Ly - 1
    End If

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
End Sub
