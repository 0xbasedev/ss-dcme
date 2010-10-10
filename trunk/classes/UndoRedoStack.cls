VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoRedoStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim undostack() As Changes
Dim redostack() As Changes

Public undocurpos As Integer
Public redocurpos As Integer

Dim maxundostack As Long

Dim parent As frmMain

Sub setParent(Main As frmMain)
10        Set parent = Main
End Sub

Private Sub Class_Initialize()
10        undocurpos = 0
20        redocurpos = 0
          
30        maxundostack = CLng(GetSetting("UndoSteps", "49"))
40        ReDim undostack(maxundostack) As Changes
50        ReDim redostack(maxundostack) As Changes
End Sub

Function GetUndoComment() As String
10        If undocurpos = 0 Then
20            GetUndoComment = ""
30        Else
40            GetUndoComment = undostack(undocurpos - 1).GetComment
50        End If
End Function

Function GetRedoComment() As String
10        If redocurpos = 0 Then
20            GetRedoComment = ""
30        Else
40            GetRedoComment = redostack(redocurpos - 1).GetComment
50        End If
End Function

Sub ResizeStack(steps As Long)

End Sub

Sub Undo()
          
10        Call Process(undostack(undocurpos - 1), True)

20        Set undostack(undocurpos - 1) = Nothing

30        If undocurpos - 1 = 0 Then
40            undocurpos = undocurpos - 1
50        Else
60            undocurpos = undocurpos - 1
70        End If
          
80        frmGeneral.UpdateToolBarButtons
End Sub

Sub Redo()
10        Call Process(redostack(redocurpos - 1), False)

20        Set redostack(redocurpos - 1) = Nothing
30        If redocurpos - 1 = 0 Then
40            redocurpos = redocurpos - 1
50        Else
60            redocurpos = redocurpos - 1
70        End If
          
80        frmGeneral.UpdateToolBarButtons
End Sub

Sub AddToUndo(ByRef undochanges As Changes, comment As String)
          
    If undochanges.changecount = 0 Then Exit Sub
    
          'adding something to undo... it means the map is changed
10        parent.mapchanged = True
          
20        Call undochanges.SetComment(comment) 'Description of UNDO action
          
      'the undo stack is full, remove the first one so we can have
      'room for the new one
30        If undocurpos = UBound(undostack) + 1 Then
40            undocurpos = UBound(undostack)
50            Call sliceUndoStack
60        End If

70        Set undostack(undocurpos) = undochanges
          
80        Set undochanges = Nothing
          
90        undocurpos = undocurpos + 1

100       frmGeneral.UpdateToolBarButtons
          
End Sub

Private Sub AddToRedo(ByRef redochanges As Changes, comment As String)

10        Call redochanges.SetComment(comment)
          
      'the undo stack is full, remove the first one so we can have
      'room for the new one
20        If redocurpos = UBound(redostack) + 1 Then
30            redocurpos = UBound(redostack)
40            Call sliceRedoStack
50        End If

60        Set redostack(redocurpos) = redochanges
70        redocurpos = redocurpos + 1
          
80        Set redochanges = Nothing
          
90        frmGeneral.UpdateToolBarButtons

End Sub

Sub sliceUndoStack()
          Dim i As Integer
10        For i = 1 To UBound(undostack)
20            Set undostack(i - 1) = undostack(i)
30        Next
40        Set undostack(maxundostack) = Nothing
End Sub

Sub sliceRedoStack()
          Dim i As Integer
10        For i = 1 To UBound(redostack)
20            Set redostack(i - 1) = redostack(i)
30        Next
40        Set redostack(maxundostack) = Nothing
End Sub

Sub Process(change As Changes, Undo As Boolean)

10        On Error GoTo Process_Error


          ' tool:X,Y,TILE:X,Y,TILE:X,Y,TILE:....
20        If Not IsObject(change) Then Exit Sub

30        frmGeneral.IsBusy("UndoRedoStack.Process") = True
          
          Dim i As Long
          Dim curtype As enumCHANGETYPE

          Dim urstr As New Changes
          Dim tilechange As typeUNDOTILECHANGE
          Dim b() As Byte

40        For i = change.ubnd To 0 Step -1
50            curtype = change.GetChangeType(i)

60            If curtype = MapTileChange Then

70                tilechange = change.GetChangeTileData(i)

80                If tilechange.tilenr = 217 Or tilechange.tilenr = 219 Or tilechange.tilenr = 220 Then
90                    Call setObject(parent, tilechange.tilenr, tilechange.X, tilechange.Y, urstr, False)
100               Else
110                   Call parent.setTile(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)
120                   Call parent.UpdateLevelTile(tilechange.X, tilechange.Y, False, False)
130               End If



140           ElseIf curtype = SelTileChange Then

150               tilechange = change.GetChangeTileData(i)
                  '            If tilechange.tilenr = 217 Or tilechange.tilenr = 219 Or tilechange.tilenr = 220 Then
                  '                Call setSelObject(parent.sel, tilechange.tilenr, tilechange.X, tilechange.Y, urstr, False)
                  '            Else
160               Call parent.sel.setSelTile(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)
                  '            End If

170           ElseIf curtype = SelAdd Then
180               tilechange = change.GetChangeTileData(i)

190               Call parent.sel.RemoveSelectionArea(tilechange.X, tilechange.X, tilechange.Y, tilechange.Y, urstr, False)

200           ElseIf curtype = SelDrop Then
210               tilechange = change.GetChangeTileData(i)

220               Call parent.sel.AddSelectionArea(tilechange.X, tilechange.X, tilechange.Y, tilechange.Y, urstr, False)
230               Call parent.setTile(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)
240               Call parent.UpdateLevelTile(tilechange.X, tilechange.Y, False)

250           ElseIf curtype = SelNew Then
260               tilechange = change.GetChangeTileData(i)
270               Call parent.sel.DeleteSelectionTile(tilechange.X, tilechange.Y, urstr)

280           ElseIf curtype = SelDelete Then
290               tilechange = change.GetChangeTileData(i)

300               Call parent.sel.CreateTileInSelection(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)

310           ElseIf curtype = SelMove Then
320               b = change.GetChangeData(i)

                  Dim dx As Integer
                  Dim dy As Integer
330               dx = -BytesToInteger(b, 0)
340               dy = -BytesToInteger(b, 2)

350               Call parent.sel.Move(dx, dy, False)
360               Call urstr.AddSelMove(dx, dy)
370               Call parent.RedrawSelection(False)

380           ElseIf curtype = SelFlip Then
390               Call parent.sel.Flip(urstr)

400           ElseIf curtype = SelMirror Then
410               Call parent.sel.Mirror(urstr)

420           ElseIf curtype = SelRotateCW Then
430               Call parent.sel.RotateCCW(urstr)

440           ElseIf curtype = SelRotateCCW Then
450               Call parent.sel.RotateCW(urstr)

460           ElseIf curtype = SelRotate180 Then
470               Call parent.sel.Rotate180(urstr)

480           ElseIf curtype = SelPaste Then

490               Call parent.sel.DeleteSelection(urstr, False, True)

500           ElseIf curtype = RegionAdd Then
510               tilechange = change.GetChangeTileData(i)
520               Call parent.Regions.RemoveFromRegion(tilechange.tilenr, tilechange.X, tilechange.Y, urstr)
              
530           ElseIf curtype = RegionRemove Then
540               tilechange = change.GetChangeTileData(i)
550               Call parent.Regions.AddToRegion(tilechange.tilenr, tilechange.X, tilechange.Y, urstr)
                  
              
560           End If


570       Next

          'objects are processed directly by the undo chunks
          '    If objcount > 0 Then
          '        For i = objcount - 1 To 0 Step -1
          '            tilechange = change.GetChangeTileData(objecttoset(i))
          '
          '            'Call setObject(parent, tilechange.tilenr, tilechange.X, tilechange.Y, urstr, False)
          '        Next
          '    End If

580       If Undo Then
590           Call AddToRedo(urstr, change.GetComment)
600       Else
610           Call AddToUndo(urstr, change.GetComment)
620       End If

630       Call parent.UpdateLevel

640       frmGeneral.IsBusy("UndoRedoStack.Process") = False
          
650       On Error GoTo 0
660       Exit Sub

Process_Error:

670       frmGeneral.IsBusy("UndoRedoStack.Process") = False
          'messagebox "Error " & Err.Number & " (" & Err.Description & ") in procedure Process of Class Module UndoRedoStack"

End Sub


Sub ResetRedo()

10        On Error GoTo ResetRedo_Error

          Dim i As Integer
20        For i = redocurpos - 1 To UBound(redostack)
30            If i >= 0 Then
40                Set redostack(i) = Nothing
50                frmGeneral.toolbartop.Buttons("Redo").Enabled = False
60                frmGeneral.mnuRedo.Enabled = False
70            End If
80        Next
90        redocurpos = 0

100       On Error GoTo 0
110       Exit Sub

ResetRedo_Error:
120       HandleError Err, "UndoRedoStack.ResetRedo"

End Sub

'Sub ResetUndo()
'
'    On Error GoTo ResetUndo_Error
'
'    Dim i As Integer
'    For i = undocurpos - 1 To UBound(undostack)
'        If i >= 0 Then
'            undostack(i) = ""
'            frmGeneral.ToolbarTop.Buttons("Undo").Enabled = False
'            frmGeneral.mnuUndo.Enabled = False
'        End If
'    Next
'    undocurpos = 0
'
'    On Error GoTo 0
'    Exit Sub
'
'ResetUndo_Error:
'
'    'messagebox "Error " & Err.Number & " (" & Err.Description & ") in procedure ResetUndo of Class Module UndoRedoStack"
'End Sub

Private Sub Class_Terminate()
10        Set parent = Nothing
          
          Dim i As Integer
20        For i = 0 To UBound(undostack)
30            Set undostack(i) = Nothing
40        Next
          
50        For i = 0 To UBound(redostack)
60            Set redostack(i) = Nothing
70        Next
End Sub
