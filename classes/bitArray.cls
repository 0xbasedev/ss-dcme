VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bitArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim b() As Byte
Dim c_NrBytes As Long

Sub ResizeArray(bytes As Long, Optional keepdata As Boolean = False)
    If keepdata Then
        ReDim Preserve b(bytes)
    Else
        ReDim b(bytes)
    End If
    c_NrBytes = bytes
End Sub

Public Property Get NrBytes()
    NrBytes = c_NrBytes
End Property

Function GetBit(bit As Long) As Boolean
    If bit \ 8 > NrBytes - 1 Then
        GetBit = False
    Else
        GetBit = b(bit \ 8) And CByte(2 ^ ((bit Mod 8)))
    End If
End Function

Function GetBitValue(bit As Long) As Double
    If bit \ 8 > NrBytes - 1 Then
        GetBitValue = 0
    Else
        GetBitValue = b(bit \ 8) And CByte(2 ^ ((bit Mod 8)))
    End If
End Function

Function GetValue(fromIdx As Long, Length As Long) As Double
    Dim i As Long
    Dim value As Double
    value = 0
    For i = fromIdx To fromIdx + Length - 1
        value = value + Abs((GetBit(i) * 2 ^ (i - fromIdx)))
    Next
    GetValue = value
End Function

Sub SwapBytes()
    Dim tmp As Byte
    Dim i As Long
    For i = 0 To UBound(b) \ 2
        tmp = b(i)
        b(i) = b(UBound(b) - 1)
        b(UBound(b) - i) = tmp
    Next
End Sub


Private Sub Class_Initialize()
    ResizeArray (2)
End Sub

Sub SetData(newB() As Byte)
    Erase b
    b = newB
    c_NrBytes = UBound(newB) + 1
End Sub

Function GetBytes() As Byte()
    GetBytes = b
End Function
