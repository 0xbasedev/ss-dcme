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
10        If keepdata Then
20            ReDim Preserve b(bytes)
30        Else
40            ReDim b(bytes)
50        End If
60        c_NrBytes = bytes
End Sub

Public Property Get NrBytes()
10        NrBytes = c_NrBytes
End Property

Function GetBit(bit As Long) As Boolean
10        If bit \ 8 > NrBytes - 1 Then
20            GetBit = False
30        Else
40            GetBit = b(bit \ 8) And CByte(2 ^ ((bit Mod 8)))
50        End If
End Function

Function GetBitValue(bit As Long) As Double
10        If bit \ 8 > NrBytes - 1 Then
20            GetBitValue = 0
30        Else
40            GetBitValue = b(bit \ 8) And CByte(2 ^ ((bit Mod 8)))
50        End If
End Function

Function GetValue(fromIdx As Long, Length As Long) As Double
          Dim i As Long
          Dim value As Double
10        value = 0
20        For i = fromIdx To fromIdx + Length - 1
30            value = value + Abs((GetBit(i) * 2 ^ (i - fromIdx)))
40        Next
50        GetValue = value
End Function

Sub SwapBytes()
          Dim tmp As Byte
          Dim i As Long
10        For i = 0 To UBound(b) \ 2
20            tmp = b(i)
30            b(i) = b(UBound(b) - 1)
40            b(UBound(b) - i) = tmp
50        Next
End Sub


Private Sub Class_Initialize()
10        ResizeArray (2)
End Sub

Sub SetData(newB() As Byte)
10        Erase b
20        b = newB
30        c_NrBytes = UBound(newB) + 1
End Sub

Function GetBytes() As Byte()
10        GetBytes = b
End Function
