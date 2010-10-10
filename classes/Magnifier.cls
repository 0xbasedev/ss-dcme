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
10        Set parent = myparent
20        If Not myparent Is Nothing Then
30            parent.currentzoom = c_zoom
40            parent.currenttilew = c_zoom * TILEW
50        End If
End Sub

Private Sub Class_Initialize()
10        c_zoom = 1
20        c_UsingPixels = False
End Sub

Public Property Get zoom() As Single
10        zoom = c_zoom
End Property

Public Property Get UsingPixels() As Boolean
10        UsingPixels = c_UsingPixels
End Property

Sub MouseDown(Button As Integer, X As Single, Y As Single)
          Dim curtilex As Integer
          Dim curtiley As Integer
10        curtilex = (parent.hScr.value + X) \ (TILEW * c_zoom)
20        curtiley = (parent.vScr.value + Y) \ (TILEW * c_zoom)
30        If (curtilex < 0 Or curtiley < 0 Or curtilex > 1023 Or curtiley > 1023) Then
40            Exit Sub
50        End If

60        If Button = vbLeftButton Then
              'when selecting, then zooming, the update for
              'selection requires not to have mouse down
              'so we need to put in before zooming, instead
              'of automaticly after
70            SharedVar.MouseDown = 0
80            Call ZoomIn(curtilex, curtiley)
90        ElseIf Button = vbRightButton Then

100           SharedVar.MouseDown = 0
110           Call ZoomOut(curtilex, curtiley)
120       End If
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

          
          
          
10        If c_zoom = 1 / 16 Then
20            Call SetZoom(1 / 4, False)
30        ElseIf c_zoom < 2 Then
40            Call SetZoom(c_zoom * 2, False)
50        Else
60            Exit Sub
70        End If


80        Call parent.SetFocusAt(tileX, tileY, parent.picpreview.width \ 2, parent.picpreview.height \ 2, False)

90        If Refresh Then
100           Call parent.UpdateLevel(False, True)
110           Call frmGeneral.UpdateToolToolbar
120       End If


          
End Sub

Sub ZoomOut(tileX As Integer, tileY As Integer, Optional Refresh As Boolean = True)
      'dont c_zoom out when reached max c_zoom out level



          
10        If parent.TestMap.isRunning Then Exit Sub

20        If c_zoom > 1 / 16 Then
30            If c_zoom = 1 / 4 Then
40                Call SetZoom(1 / 16, False)
50            Else
60                Call SetZoom(c_zoom / 2, False)
70            End If
80        Else
90            Exit Sub
100       End If
          
110       Call parent.SetFocusAt(tileX, tileY, parent.picpreview.width \ 2, parent.picpreview.height \ 2, False)

120       If Refresh Then
130           Call parent.UpdateLevel(False, True)
140           Call frmGeneral.UpdateToolToolbar
150       End If

End Sub

Sub SetZoom(newzoom As Single, Refresh As Boolean)
10        c_zoom = newzoom
          
20        c_UsingPixels = (c_zoom = 1 / 16)
          
          
          
30        With parent
                  
              Dim curtilew As Long
40            curtilew = c_zoom * TILEW
              
50            .currentzoom = c_zoom
60            .currenttilew = curtilew
              
70            .picempty.width = curtilew
80            .picempty.height = curtilew
90            .pictrans.width = curtilew * 6
100           .pictrans.height = curtilew * 6
              
110           Call .UpdateScrollbars(False)
              'recalculate the scrollbars


              'build the zoomed tileset:
          
              'make it the correct size
120           .piczoomtileset.width = .pictileset.width * c_zoom
130           .piczoomtileset.height = .pictileset.height * c_zoom
140           .picHighlightZoomTileset.width = .pictileset.width * c_zoom
150           .picHighlightZoomTileset.height = .pictileset.height * c_zoom

              'stretchblt the normal tileset to it
          
160           SetStretchBltMode .piczoomtileset.hDC, HALFTONE
170           StretchBlt .piczoomtileset.hDC, 0, 0, .piczoomtileset.width, .piczoomtileset.height, .pictileset.hDC, 0, 0, .pictileset.width, .pictileset.height, vbSrcCopy
180           .piczoomtileset.Refresh

190           SetStretchBltMode .picHighlightZoomTileset.hDC, HALFTONE
200           StretchBlt .picHighlightZoomTileset.hDC, 0, 0, .piczoomtileset.width, .piczoomtileset.height, .pichighlightTileset.hDC, 0, 0, .pichighlightTileset.width, .pichighlightTileset.height, vbSrcCopy
210           .picHighlightZoomTileset.Refresh
          
          Call parent.lvz.buildAllZoomImages(c_zoom)
          
220           If Refresh Then
230               Call .UpdateLevel(False, True)
240               Call frmGeneral.UpdateToolToolbar
250           End If
              
260       End With
          
270       frmGeneral.toolbartop.Buttons("ZoomIn").Enabled = (c_zoom < 2)
280       frmGeneral.toolbartop.Buttons("ZoomOut").Enabled = (c_zoom > 1 / 16)


End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
End Sub
