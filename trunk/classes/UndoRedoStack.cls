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
    Set parent = Main
End Sub

Private Sub Class_Initialize()
    undocurpos = 0
    redocurpos = 0
    
    maxundostack = CLng(GetSetting("UndoSteps", "49"))
    ReDim undostack(maxundostack) As Changes
    ReDim redostack(maxundostack) As Changes
End Sub

Function GetUndoComment() As String
    If undocurpos = 0 Then
        GetUndoComment = ""
    Else
        GetUndoComment = undostack(undocurpos - 1).GetComment
    End If
End Function

Function GetRedoComment() As String
    If redocurpos = 0 Then
        GetRedoComment = ""
    Else
        GetRedoComment = redostack(redocurpos - 1).GetComment
    End If
End Function

Sub ResizeStack(steps As Long)

End Sub

Sub Undo()
    
    Call Process(undostack(undocurpos - 1), True)

    Set undostack(undocurpos - 1) = Nothing

    If undocurpos - 1 = 0 Then
        undocurpos = undocurpos - 1
    Else
        undocurpos = undocurpos - 1
    End If
    
    frmGeneral.UpdateToolBarButtons
End Sub

Sub Redo()
    Call Process(redostack(redocurpos - 1), False)

    Set redostack(redocurpos - 1) = Nothing
    If redocurpos - 1 = 0 Then
        redocurpos = redocurpos - 1
    Else
        redocurpos = redocurpos - 1
    End If
    
    frmGeneral.UpdateToolBarButtons
End Sub

Sub AddToUndo(ByRef undochanges As Changes, comment As String)
    
    If undochanges.changecount = 0 Then Exit Sub
    
    'adding something to undo... it means the map is changed
    parent.mapchanged = True
    
    Call undochanges.SetComment(comment) 'Description of UNDO action
    
'the undo stack is full, remove the first one so we can have
'room for the new one
    If undocurpos = UBound(undostack) + 1 Then
        undocurpos = UBound(undostack)
        Call sliceUndoStack
    End If

    Set undostack(undocurpos) = undochanges
    
    Set undochanges = Nothing
    
    undocurpos = undocurpos + 1

    frmGeneral.UpdateToolBarButtons
    
End Sub

Private Sub AddToRedo(ByRef redochanges As Changes, comment As String)

    Call redochanges.SetComment(comment)
    
'the undo stack is full, remove the first one so we can have
'room for the new one
    If redocurpos = UBound(redostack) + 1 Then
        redocurpos = UBound(redostack)
        Call sliceRedoStack
    End If

    Set redostack(redocurpos) = redochanges
    redocurpos = redocurpos + 1
    
    Set redochanges = Nothing
    
    frmGeneral.UpdateToolBarButtons

End Sub

Sub sliceUndoStack()
    Dim i As Integer
    For i = 1 To UBound(undostack)
        Set undostack(i - 1) = undostack(i)
    Next
    Set undostack(maxundostack) = Nothing
End Sub

Sub sliceRedoStack()
    Dim i As Integer
    For i = 1 To UBound(redostack)
        Set redostack(i - 1) = redostack(i)
    Next
    Set redostack(maxundostack) = Nothing
End Sub

Sub Process(Change As Changes, Undo As Boolean)

    On Error GoTo Process_Error


    ' tool:X,Y,TILE:X,Y,TILE:X,Y,TILE:....
    If Not IsObject(Change) Then Exit Sub

    frmGeneral.IsBusy("UndoRedoStack.Process") = True
    
    Dim i As Long
    Dim curtype As enumCHANGETYPE

    Dim urstr As New Changes
    Dim tilechange As typeUNDOTILECHANGE
    Dim b() As Byte

    For i = Change.ubnd To 0 Step -1
        curtype = Change.GetChangeType(i)

        If curtype = MapTileChange Then

            tilechange = Change.GetChangeTileData(i)

            If tilechange.tilenr = 217 Or tilechange.tilenr = 219 Or tilechange.tilenr = 220 Then
                Call setObject(parent, tilechange.tilenr, tilechange.X, tilechange.Y, urstr, False)
            Else
                Call parent.setTile(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)
                Call parent.UpdateLevelTile(tilechange.X, tilechange.Y, False, False)
            End If



        ElseIf curtype = SelTileChange Then

            tilechange = Change.GetChangeTileData(i)
            '            If tilechange.tilenr = 217 Or tilechange.tilenr = 219 Or tilechange.tilenr = 220 Then
            '                Call setSelObject(parent.sel, tilechange.tilenr, tilechange.X, tilechange.Y, urstr, False)
            '            Else
            Call parent.sel.setSelTile(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)
            '            End If

        ElseIf curtype = SelAdd Then
            tilechange = Change.GetChangeTileData(i)

            Call parent.sel.RemoveSelectionArea(tilechange.X, tilechange.X, tilechange.Y, tilechange.Y, urstr, False)

        ElseIf curtype = SelDrop Then
            tilechange = Change.GetChangeTileData(i)

            Call parent.sel.AddSelectionArea(tilechange.X, tilechange.X, tilechange.Y, tilechange.Y, urstr, False)
            Call parent.setTile(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)
            Call parent.UpdateLevelTile(tilechange.X, tilechange.Y, False)

        ElseIf curtype = SelNew Then
            tilechange = Change.GetChangeTileData(i)
            Call parent.sel.DeleteSelectionTile(tilechange.X, tilechange.Y, urstr)

        ElseIf curtype = SelDelete Then
            tilechange = Change.GetChangeTileData(i)

            Call parent.sel.CreateTileInSelection(tilechange.X, tilechange.Y, tilechange.tilenr, urstr)

        ElseIf curtype = SelMove Then
            b = Change.GetChangeData(i)

            Dim dx As Integer
            Dim dy As Integer
            dx = -BytesToInteger(b, 0)
            dy = -BytesToInteger(b, 2)

            Call parent.sel.Move(dx, dy, False)
            Call urstr.AddSelMove(dx, dy)
            Call parent.RedrawSelection(False)

        ElseIf curtype = SelFlip Then
            Call parent.sel.Flip(urstr)

        ElseIf curtype = SelMirror Then
            Call parent.sel.Mirror(urstr)

        ElseIf curtype = SelRotateCW Then
            Call parent.sel.RotateCCW(urstr)

        ElseIf curtype = SelRotateCCW Then
            Call parent.sel.RotateCW(urstr)

        ElseIf curtype = SelRotate180 Then
            Call parent.sel.Rotate180(urstr)

        ElseIf curtype = SelPaste Then

            Call parent.sel.DeleteSelection(urstr, False, True)

        ElseIf curtype = RegionAdd Then
            tilechange = Change.GetChangeTileData(i)
            Call parent.Regions.RemoveFromRegion(tilechange.tilenr, tilechange.X, tilechange.Y, urstr)
        
        ElseIf curtype = RegionRemove Then
            tilechange = Change.GetChangeTileData(i)
            Call parent.Regions.AddToRegion(tilechange.tilenr, tilechange.X, tilechange.Y, urstr)
            
        
        End If


    Next

    'objects are processed directly by the undo chunks
    '    If objcount > 0 Then
    '        For i = objcount - 1 To 0 Step -1
    '            tilechange = change.GetChangeTileData(objecttoset(i))
    '
    '            'Call setObject(parent, tilechange.tilenr, tilechange.X, tilechange.Y, urstr, False)
    '        Next
    '    End If

    If Undo Then
        Call AddToRedo(urstr, Change.GetComment)
    Else
        Call AddToUndo(urstr, Change.GetComment)
    End If

    Call parent.UpdateLevel

    frmGeneral.IsBusy("UndoRedoStack.Process") = False
    
    On Error GoTo 0
    Exit Sub

Process_Error:

    frmGeneral.IsBusy("UndoRedoStack.Process") = False
    'messagebox "Error " & Err.Number & " (" & Err.Description & ") in procedure Process of Class Module UndoRedoStack"

End Sub


Sub ResetRedo()

    On Error GoTo ResetRedo_Error

    Dim i As Integer
    For i = redocurpos - 1 To UBound(redostack)
        If i >= 0 Then
            Set redostack(i) = Nothing
            frmGeneral.ToolbarTop.Buttons("Redo").Enabled = False
            frmGeneral.mnuRedo.Enabled = False
        End If
    Next
    redocurpos = 0

    On Error GoTo 0
    Exit Sub

ResetRedo_Error:
    HandleError Err, "UndoRedoStack.ResetRedo"

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
    Set parent = Nothing
    
    Dim i As Integer
    For i = 0 To UBound(undostack)
        Set undostack(i) = Nothing
    Next
    
    For i = 0 To UBound(redostack)
        Set redostack(i) = Nothing
    Next
End Sub
