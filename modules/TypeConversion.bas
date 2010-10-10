Attribute VB_Name = "TypeConversion"
Option Explicit


Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

'Converts 4 bytes to a double
Function bytesToDouble(b() As Byte) As Double
          Dim i As Integer
          Dim value As Double
10        For i = 0 To 3
20            value = value + b(i) * 256 ^ i
30        Next
40        bytesToDouble = value
End Function

Function SignedTo12bits(ByVal value As Integer) As Integer
10        If value < 0 Then
20            SignedTo12bits = 4096 + value
30        Else
40            SignedTo12bits = value
50        End If
End Function

Function SignedFrom12bits(ByVal value As Integer) As Integer
10        If value > 2047 Then
20            SignedFrom12bits = value - 4096
30        Else
40            SignedFrom12bits = value
50        End If
End Function


Function UnsignedToLong(ByVal value As Double) As Long
10        If value < 0 Or value >= OFFSET_4 Then Error 6    ' Overflow
20        If value <= MAXINT_4 Then
30            UnsignedToLong = value
40        Else
50            UnsignedToLong = value - OFFSET_4
60        End If
End Function

Function LongToUnsigned(ByVal value As Long) As Double
10        If value < 0 Then
20            LongToUnsigned = value + OFFSET_4
30        Else
40            LongToUnsigned = value
50        End If
End Function

Function UnsignedToInteger(ByVal value As Long) As Integer
10        If value < 0 Or value >= OFFSET_2 Then Error 6    ' Overflow
20        If value <= MAXINT_2 Then
30            UnsignedToInteger = value
40        Else
50            UnsignedToInteger = value - OFFSET_2
60        End If
End Function

Function IntegerToUnsigned(ByVal value As Integer) As Long
10        If value < 0 Then
20            IntegerToUnsigned = value + OFFSET_2
30        Else
40            IntegerToUnsigned = value
50        End If
End Function

Function IntegerToBytes(ByVal value As Long) As Byte()
          Dim b() As Byte
10        ReDim b(1) As Byte
          Dim i As Integer
          
20        If value <= OFFSET_2 Then
30            For i = 1 To 2
40                b(i - 1) = value Mod 256
50                value = value \ 256
60            Next
70        Else
80            b(0) = 0
90            b(1) = 0
100       End If
          
110       IntegerToBytes = b
End Function


'Converts 4 bytes to a long
Function bytesToLong(b() As Byte) As Long
          Dim i As Integer
          Dim value As Double
10        For i = 0 To 3
20            value = value + b(i) * 256 ^ i
30        Next
40        bytesToLong = UnsignedToLong(value)
End Function

Function LongtoBytes(ByVal value As Long) As Byte()
          Dim b() As Byte
10        ReDim b(3) As Byte
          Dim i As Integer

20        For i = 1 To 4
30            b(i - 1) = value Mod 256
40            value = value \ 256
50        Next
60        LongtoBytes = b
End Function

'Converts a double to a 4-bytes array (u32)
Function DoubletoBytes(ByVal value As Double) As Byte()
          
          Dim b() As Byte
10        ReDim b(3) As Byte
          Dim i As Integer

20        If value > OFFSET_4 Then
30            For i = 0 To 3
40                b(i) = 0
50            Next
60        Else
70            For i = 1 To 4
80                b(i - 1) = value Mod 256
90                value = value \ 256
100           Next
110       End If
120       DoubletoBytes = b
End Function

Function bytesToString(b() As Byte) As String
          Dim i As Integer
          Dim str As String

10        For i = 0 To UBound(b)
20            If b(i) <> 0 Then
30                str = str & Chr(b(i))
40            End If
50        Next

60        bytesToString = str
End Function

Function stringToBytes(str As String) As Byte()
          Dim b() As Byte
10        ReDim b(Len(str) - 1) As Byte

          Dim i As Integer
20        For i = 1 To Len(str)
30            b(i - 1) = Asc(Mid$(str, i, 1))
40        Next

50        stringToBytes = b
End Function


Function IntegerToByteArray(ByVal intvalue As Long) As Byte()

      'Example:
      'dim bytArr() as Byte
      'dim iCtr as Integer
      'bytArr = LongToByteArray(90121)
      'For iCtr = 0 to Ubound(bytArr)
      'Debug.Print bytArr(iCtr)
      'Next
      '******************************************
          Dim ByteArray(0 To 1) As Byte
10        CopyMemory ByteArray(0), ByVal VarPtr(intvalue), Len(intvalue)
20        IntegerToByteArray = ByteArray

End Function

Function LongToByteArray(ByVal lng As Long) As Byte()

      'Example:
      'dim bytArr() as Byte
      'dim iCtr as Integer
      'bytArr = LongToByteArray(90121)
      'For iCtr = 0 to Ubound(bytArr)
      'Debug.Print bytArr(iCtr)
      'Next
      '******************************************
          Dim ByteArray(0 To 3) As Byte
10        CopyMemory ByteArray(0), ByVal VarPtr(lng), Len(lng)
20        LongToByteArray = ByteArray

End Function

Function BytesToInteger(ByteArray() As Byte, FirstIndex As Long) As Integer
          Dim value As Long

10        value = ByteArray(FirstIndex) + ByteArray(FirstIndex + 1) * 256&

20        BytesToInteger = UnsignedToInteger(value)

End Function


Function BytesToNumEx(ByteArray() As Byte, StartRec As Long, _
                      EndRec As Long, UnSigned As Boolean) As Double
      ' ###################################################
      ' Author                : Imran Zaheer
      ' Contact               : imraanz@mail.com
      ' Date                  : January 2000
      ' Function BytesToNumEx : Convertes the specified byte array
      '                         into the corresponding Integer or Long
      '                         or any signed/unsigned
      '                        ;(non-float) data type.
      '
      ' * BYTES : LIKE NUMBERS(Integer/Long etc.) STORED IN A
      ' * BINARY FILE

      ' Parameters :
      '  (All parameters are reuuired: No Optional)
      '     ByteArray() : byte array containg a number in byte format
      '  StartRec    : specify the starting array record within the
      ' array
      '     EndRec      : specify the end array record within the array
      '     UnSigned    : when False process bytes for both -ve and
      '                   +ve values.
      '                   when true only process the bytes for +ve
      '                   values.
      '
      ' Note: If both "StartRec" and "EndRec" Parameters are zero,
      '       then the complete array will be processed.
      '
      ' Example Calls :
      '      dim myArray(1 To 4) as byte
      '      dim myVar1 as Integer
      '      dim myVar2 as Long
      '
      '      myArray(1) = 255
      '      myArray(2) = 127
      '      myVar1 = BytesToNumEx(myArray(), 1, 2, False)
      '  after execution of above statement myVar1 will be 32767
      '
      '      myArray(1) = 0
      '      myArray(2) = 0
      '      myArray(3) = 0
      '      myArray(4) = 128
      '      myVar2 = BytesToNumEx(myArray(), 1, 4, False)
      '  after execution of above statement myVar2 will be -2147483648
      '
      '
      '####################################################
10        On Error GoTo ErrorHandler
          Dim i As Integer
          Dim lng256 As Double
          Dim lngReturn As Double

20        lng256 = 1
30        lngReturn = 0

40        If EndRec < 1 Then
50            EndRec = UBound(ByteArray)
60        End If

70        If StartRec > EndRec Or StartRec < 0 Then
80            MessageBox _
                      "Start record can not be greater then End record...!", _
                      vbInformation
90            BytesToNumEx = -1
100           Exit Function
110       End If

120       lngReturn = lngReturn + (ByteArray(StartRec))
130       For i = (StartRec + 1) To EndRec
140           lng256 = lng256 * 256
150           If i < EndRec Then
160               lngReturn = lngReturn + (ByteArray(i) * lng256)
170           Else
                  ' if -ve

180               If ByteArray(i) > 127 And UnSigned = False Then
190                   lngReturn = (lngReturn + ((ByteArray(i) - 256) _
                                                * lng256))
200               Else
210                   lngReturn = lngReturn + (ByteArray(i) * lng256)
220               End If
230           End If
240       Next i

250       BytesToNumEx = lngReturn
ErrorHandler:
End Function
