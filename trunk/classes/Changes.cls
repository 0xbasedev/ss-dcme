VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Changes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" _
                               Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal _
                                                                                         Length As Long)
'In SharedVar module
'Type Change
'    ChangeType As enumCHANGETYPE
'    ChangeData() As Byte
'End Type
'
'Enum enumCHANGETYPE
'    'map-related
'    MapTileChange   'map tile changed
'
'    SelTileChange   'tile changed within selection
'    SelAdd          'map tile added to selection
'    SelDrop         'tile dropped from selection to map
'    SelMove         'selection moved
'    SelDelete       'tile cleared from sel without being dropped
'    SelNew          'new tile created in selection (pasting, for example)
'
'    'region-related
'    RegionAdd       'tile added to region
'    RegionRemove    'tile removed from region
'    RegionNew       'new region created
'    RegionDelete    'region deleted
'    RegionRename    'region renamed
'    RegionProperties 'region properties changed
'
'    'lvz-related
'
'    'misc
'End Enum

Dim changecomment As String
Dim Change() As typeUNDOACTION
Dim curpos As Long


Public Property Get changecount() As Long
    changecount = curpos
End Property

Function ubnd() As Long
    If curpos > 0 Then
        ubnd = curpos - 1
    Else
        ubnd = 0
    End If
End Function

Private Sub Class_Initialize()
    ReDim Change(100)
    curpos = 0
End Sub

'used for:      (what to set 'tile' to)
'MapTileChange  (old tilenr)
'SelTileChange  (old sel tilenr)
'SelAdd         (old tilenr = 0)
'SelDrop        (old map tilenr)
'SelDelete      (sel tilenr)
'SelNew         (new sel tilenr)
'RegionAdd      (regionID)
'RegionRemove   (regionID)
Sub AddTileChange(ByRef ChangeType As enumCHANGETYPE, ByRef X As Integer, ByRef Y As Integer, ByRef tile As Integer)

    Dim b(4) As Byte

    '3 bytes: X/Y
    '2 bytes: oldtilenr

    AppendCoordinatesToBytes b, 0, X, Y

    AppendIntegerToBytes b, 3, tile

    Call AddChange(ChangeType, b)
End Sub

Sub AddSelMove(dx As Integer, dy As Integer)
    Dim b(3) As Byte
    'first 2 bytes, dx (signed)
    ' last 2 bytes, dy (signed)
    AppendIntegerToBytes b, 0, dx
    AppendIntegerToBytes b, 2, dy

    Call AddChange(SelMove, b)
End Sub

Sub AddSelFlip(horizontal As Boolean)
    Dim b(0) As Byte
    If horizontal Then
        Call AddChange(SelMirror, b)
    Else
        Call AddChange(SelFlip, b)
    End If
End Sub

Sub AddSelRotate(angle As Integer)
    Dim b(0) As Byte
    If angle = 90 Then
        Call AddChange(SelRotateCW, b)  'selection was rotated 90 degrees
    ElseIf angle = 270 Then
        Call AddChange(SelRotateCCW, b)    'selection was rotated 270 degrees
    ElseIf angle = 180 Then
        Call AddChange(SelRotate180, b)    'selection was rotated 180 degrees
    End If
End Sub

Sub AddSelPaste()
    Dim b(0) As Byte
    Call AddChange(SelPaste, b)
End Sub

Sub AddChange(ChangeType As enumCHANGETYPE, ChangeData() As Byte)
    If curpos > UBound(Change) Then
        ReDim Preserve Change(UBound(Change) + 1000)
    End If

    Change(curpos).ChgType = ChangeType
    Change(curpos).ChgData = ChangeData

    curpos = curpos + 1

End Sub


Friend Function GetChangeType(Index As Long) As enumCHANGETYPE
    GetChangeType = Change(Index).ChgType
End Function

Friend Function GetChangeData(Index As Long) As Byte()
    GetChangeData = Change(Index).ChgData
End Function

Friend Function GetChangeTileData(Index As Long) As typeUNDOTILECHANGE

    GetChangeTileData.X = Change(Index).ChgData(0) + 256 * (Change(Index).ChgData(1) Mod 16)
    GetChangeTileData.Y = Change(Index).ChgData(1) \ 16 + 16 * Change(Index).ChgData(2)

    GetChangeTileData.tilenr = BytesToInteger(Change(Index).ChgData(), 3)

End Function


Sub AppendCoordinatesToBytes(ByRef ByteArray() As Byte, start As Long, ByRef X As Integer, ByRef Y As Integer)
    ByteArray(start) = X Mod 256    'bottom 8 of x
    ByteArray(start + 2) = Y \ 16  'top 8 of y
'calculate b(1) out of y, but add the remainder from the x
'equation because when y is calculated, it's rounded
'down and the remainder is lost
    ByteArray(start + 1) = (Y - 16 * ByteArray(start + 2)) * 16 + (X - ByteArray(start)) / 256
End Sub

Sub AppendIntegerToBytes(ByRef ByteArray() As Byte, start As Long, intvalue As Integer)
    CopyMemory ByteArray(start), ByVal VarPtr(intvalue), Len(intvalue)
End Sub

Sub AppendBytes(ByRef ByteArray() As Byte, start As Long, Data() As Byte)
    CopyMemory ByteArray(start), ByVal VarPtr(Data(0)), UBound(Data) + 1
End Sub

Sub SetComment(comment As String)
    changecomment = comment
End Sub

Function GetComment() As String
    GetComment = changecomment
End Function
