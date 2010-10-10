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
Dim change() As typeUNDOACTION
Dim curpos As Long


Public Property Get changecount() As Long
    changecount = curpos
End Property

Function ubnd() As Long
10        If curpos > 0 Then
20            ubnd = curpos - 1
30        Else
40            ubnd = 0
50        End If
End Function

Private Sub Class_Initialize()
10        ReDim change(100)
20        curpos = 0
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

10        AppendCoordinatesToBytes b, 0, X, Y

20        AppendIntegerToBytes b, 3, tile

30        Call AddChange(ChangeType, b)
End Sub

Sub AddSelMove(dx As Integer, dy As Integer)
          Dim b(3) As Byte
          'first 2 bytes, dx (signed)
          ' last 2 bytes, dy (signed)
10        AppendIntegerToBytes b, 0, dx
20        AppendIntegerToBytes b, 2, dy

30        Call AddChange(SelMove, b)
End Sub

Sub AddSelFlip(horizontal As Boolean)
          Dim b(0) As Byte
10        If horizontal Then
20            Call AddChange(SelMirror, b)
30        Else
40            Call AddChange(SelFlip, b)
50        End If
End Sub

Sub AddSelRotate(angle As Integer)
          Dim b(0) As Byte
10        If angle = 90 Then
20            Call AddChange(SelRotateCW, b)  'selection was rotated 90 degrees
30        ElseIf angle = 270 Then
40            Call AddChange(SelRotateCCW, b)    'selection was rotated 270 degrees
50        ElseIf angle = 180 Then
60            Call AddChange(SelRotate180, b)    'selection was rotated 180 degrees
70        End If
End Sub

Sub AddSelPaste()
          Dim b(0) As Byte
10        Call AddChange(SelPaste, b)
End Sub

Sub AddChange(ChangeType As enumCHANGETYPE, ChangeData() As Byte)
10        If curpos > UBound(change) Then
20            ReDim Preserve change(UBound(change) + 1000)
30        End If

40        change(curpos).ChgType = ChangeType
50        change(curpos).ChgData = ChangeData

60        curpos = curpos + 1

End Sub


Friend Function GetChangeType(Index As Long) As enumCHANGETYPE
10        GetChangeType = change(Index).ChgType
End Function

Friend Function GetChangeData(Index As Long) As Byte()
10        GetChangeData = change(Index).ChgData
End Function

Friend Function GetChangeTileData(Index As Long) As typeUNDOTILECHANGE

10        GetChangeTileData.X = change(Index).ChgData(0) + 256 * (change(Index).ChgData(1) Mod 16)
20        GetChangeTileData.Y = change(Index).ChgData(1) \ 16 + 16 * change(Index).ChgData(2)

30        GetChangeTileData.tilenr = BytesToInteger(change(Index).ChgData(), 3)

End Function


Sub AppendCoordinatesToBytes(ByRef ByteArray() As Byte, start As Long, ByRef X As Integer, ByRef Y As Integer)
10        ByteArray(start) = X Mod 256    'bottom 8 of x
20        ByteArray(start + 2) = Y \ 16  'top 8 of y
      'calculate b(1) out of y, but add the remainder from the x
      'equation because when y is calculated, it's rounded
      'down and the remainder is lost
30        ByteArray(start + 1) = (Y - 16 * ByteArray(start + 2)) * 16 + (X - ByteArray(start)) / 256
End Sub

Sub AppendIntegerToBytes(ByRef ByteArray() As Byte, start As Long, intvalue As Integer)
10        CopyMemory ByteArray(start), ByVal VarPtr(intvalue), Len(intvalue)
End Sub

Sub AppendBytes(ByRef ByteArray() As Byte, start As Long, Data() As Byte)
10        CopyMemory ByteArray(start), ByVal VarPtr(Data(0)), UBound(Data) + 1
End Sub

Sub SetComment(comment As String)
10        changecomment = comment
End Sub

Function GetComment() As String
10        GetComment = changecomment
End Function
