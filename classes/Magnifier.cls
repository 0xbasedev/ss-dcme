VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Magnifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim c_zoom As Single
Dim c_UsingPixels As Boolean
Dim parent As frmMain

Sub setParent(ByRef myparent As frmMain)
    Set parent = myparent
    If Not myparent Is Nothing Then
        parent.currentzoom = c_zoom
        parent.currenttilew = c_zoom * TILEW
    End If
End Sub

Private Sub Class_Initialize()
    c_zoom = 1
    c_UsingPixels = False
End Sub

Public Property Get zoom() As Single
    zoom = c_zoom
End Property

Public Property Get UsingPixels() As Boolean
    UsingPixels = c_UsingPixels
End Property

Sub MouseDown(Button As Integer, X As Single, Y As Single)
    Dim curtilex As Integer
    Dim curtiley As Integer
    curtilex = (parent.Hscr.value + X) \ (TILEW * c_zoom)
    curtiley = (parent.Vscr.value + Y) \ (TILEW * c_zoom)
    If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
        Exit Sub
    End If

    If Button = vbLeftButton Then
        'when selecting, then zooming, the update for
        'selection requires not to have mouse down
        'so we need to put in before zooming, instead
        'of automaticly after
        SharedVar.MouseDown = 0
        Call ZoomIn(curtilex, curtiley)
    ElseIf Button = vbRightButton Then

        SharedVar.MouseDown = 0
        Call ZoomOut(curtilex, curtiley)
    End If
End Sub

Sub MouseMove()

End Sub

Sub MouseUp()
'nothing to do on mouseup
End Sub

'Private Function CheckScrollbarH(tileX As Integer) As Long
'    Dim newH As Long
'    Dim curtilew As Long
'
'    curtilew = c_zoom * TILEW
'
'    If curtilew * tileX - (parent.picPreview.width \ 2) < 0 Then
'        newH = 0
'    ElseIf curtilew * tileX - (parent.picPreview.width \ 2) > parent.Hscr.Max Then
'        newH = parent.Hscr.Max
'    Else
'        newH = curtilew * tileX - (parent.picPreview.width \ 2)
'    End If
'
'    'make a correction so the value is always divisible
'    CheckScrollbarH = newH - (newH Mod curtilew)
'
'End Function
'
'Private Function CheckScrollbarV(tileY As Integer) As Long
'    Dim newV As Long
'
'    Dim curtilew As Long
'
'    curtilew = c_zoom * TILEH
'
'    If curtilew * tileY - (parent.picPreview.height \ 2) < 0 Then
'        newV = 0
'    ElseIf curtilew * tileY - (parent.picPreview.height \ 2) > parent.Vscr.Max Then
'        newV = parent.Vscr.Max
'    Else
'        newV = curtilew * tileY - (parent.picPreview.height \ 2)
'    End If
'
'    CheckScrollbarV = newV - (newV Mod curtilew)
'End Function


Sub ZoomIn(tileX As Integer, tileY As Integer, Optional Refresh As Boolean = True)

'dont zoom in more than 2

    
    
    
    If c_zoom = 1 / 16 Then
        Call SetZoom(1 / 4, False)
    ElseIf c_zoom < 2 Then
        Call SetZoom(c_zoom * 2, False)
    Else
        Exit Sub
    End If


    Call parent.SetFocusAt(tileX, tileY, parent.picPreview.width \ 2, parent.picPreview.height \ 2, False)

    If Refresh Then
        Call parent.UpdateLevel(False, True)
        Call frmGeneral.UpdateToolToolbar
    End If


    
End Sub

Sub ZoomOut(tileX As Integer, tileY As Integer, Optional Refresh As Boolean = True)
'dont c_zoom out when reached max c_zoom out level



    
    If parent.TestMap.isRunning Then Exit Sub

    If c_zoom > 1 / 16 Then
        If c_zoom = 1 / 4 Then
            Call SetZoom(1 / 16, False)
        Else
            Call SetZoom(c_zoom / 2, False)
        End If
    Else
        Exit Sub
    End If
    
    Call parent.SetFocusAt(tileX, tileY, parent.picPreview.width \ 2, parent.picPreview.height \ 2, False)

    If Refresh Then
        Call parent.UpdateLevel(False, True)
        Call frmGeneral.UpdateToolToolbar
    End If

End Sub

Sub SetZoom(newzoom As Single, Refresh As Boolean)
    c_zoom = newzoom
    
    c_UsingPixels = (c_zoom = 1 / 16)
    
    
    
    With parent
            
        Dim curtilew As Long
        curtilew = c_zoom * TILEW
        
        .currentzoom = c_zoom
        .currenttilew = curtilew
        
        .picempty.width = curtilew
        .picempty.height = curtilew
        .pictrans.width = curtilew * 6
        .pictrans.height = curtilew * 6
        
        Call .UpdateScrollbars(False)
        'recalculate the scrollbars


        'build the zoomed tileset:
    
        'make it the correct size
        .piczoomtileset.width = .pictileset.width * c_zoom
        .piczoomtileset.height = .pictileset.height * c_zoom
        .picHighlightZoomTileset.width = .pictileset.width * c_zoom
        .picHighlightZoomTileset.height = .pictileset.height * c_zoom

        'stretchblt the normal tileset to it
    
        SetStretchBltMode .piczoomtileset.hDC, HALFTONE
        StretchBlt .piczoomtileset.hDC, 0, 0, .piczoomtileset.width, .piczoomtileset.height, .pictileset.hDC, 0, 0, .pictileset.width, .pictileset.height, vbSrcCopy
        .piczoomtileset.Refresh

        SetStretchBltMode .picHighlightZoomTileset.hDC, HALFTONE
        StretchBlt .picHighlightZoomTileset.hDC, 0, 0, .piczoomtileset.width, .piczoomtileset.height, .pichighlightTileset.hDC, 0, 0, .pichighlightTileset.width, .pichighlightTileset.height, vbSrcCopy
        .picHighlightZoomTileset.Refresh
    
    Call parent.lvz.buildAllZoomImages(c_zoom)
    
        If Refresh Then
            Call .UpdateLevel(False, True)
            Call frmGeneral.UpdateToolToolbar
        End If
        
    End With
    
    frmGeneral.ToolbarTop.Buttons("ZoomIn").Enabled = (c_zoom < 2)
    frmGeneral.ToolbarTop.Buttons("ZoomOut").Enabled = (c_zoom > 1 / 16)


End Sub

Private Sub Class_Terminate()
    Set parent = Nothing
End Sub
