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
    For i = 0 To 3
        value = value + b(i) * 256 ^ i
    Next
    bytesToDouble = value
End Function

Function SignedTo12bits(ByVal value As Integer) As Integer
    If value < 0 Then
        SignedTo12bits = 4096 + value
    Else
        SignedTo12bits = value
    End If
End Function

Function SignedFrom12bits(ByVal value As Integer) As Integer
    If value > 2047 Then
        SignedFrom12bits = value - 4096
    Else
        SignedFrom12bits = value
    End If
End Function


Function UnsignedToLong(ByVal value As Double) As Long
    If value < 0 Or value >= OFFSET_4 Then Error 6    ' Overflow
    If value <= MAXINT_4 Then
        UnsignedToLong = value
    Else
        UnsignedToLong = value - OFFSET_4
    End If
End Function

Function LongToUnsigned(ByVal value As Long) As Double
    If value < 0 Then
        LongToUnsigned = value + OFFSET_4
    Else
        LongToUnsigned = value
    End If
End Function

Function UnsignedToInteger(ByVal value As Long) As Integer
    If value < 0 Or value >= OFFSET_2 Then Error 6    ' Overflow
    If value <= MAXINT_2 Then
        UnsignedToInteger = value
    Else
        UnsignedToInteger = value - OFFSET_2
    End If
End Function

Function IntegerToUnsigned(ByVal value As Integer) As Long
    If value < 0 Then
        IntegerToUnsigned = value + OFFSET_2
    Else
        IntegerToUnsigned = value
    End If
End Function

Function IntegerToBytes(ByVal value As Long) As Byte()
    Dim b() As Byte
    ReDim b(1) As Byte
    Dim i As Integer
    
    If value <= OFFSET_2 Then
        For i = 1 To 2
            b(i - 1) = value Mod 256
            value = value \ 256
        Next
    Else
        b(0) = 0
        b(1) = 0
    End If
    
    IntegerToBytes = b
End Function


'Converts 4 bytes to a long
Function bytesToLong(b() As Byte) As Long
    Dim i As Integer
    Dim value As Double
    For i = 0 To 3
        value = value + b(i) * 256 ^ i
    Next
    bytesToLong = UnsignedToLong(value)
End Function

Function LongtoBytes(ByVal value As Long) As Byte()
    Dim b() As Byte
    ReDim b(3) As Byte
    Dim i As Integer

    For i = 1 To 4
        b(i - 1) = value Mod 256
        value = value \ 256
    Next
    LongtoBytes = b
End Function

'Converts a double to a 4-bytes array (u32)
Function DoubletoBytes(ByVal value As Double) As Byte()
    
    Dim b() As Byte
    ReDim b(3) As Byte
    Dim i As Integer

    If value > OFFSET_4 Then
        For i = 0 To 3
            b(i) = 0
        Next
    Else
        For i = 1 To 4
            b(i - 1) = value Mod 256
            value = value \ 256
        Next
    End If
    DoubletoBytes = b
End Function

Function bytesToString(b() As Byte) As String
    Dim i As Integer
    Dim str As String

    For i = 0 To UBound(b)
        If b(i) <> 0 Then
            str = str & Chr(b(i))
        End If
    Next

    bytesToString = str
End Function

Function stringToBytes(str As String) As Byte()
    Dim b() As Byte
    ReDim b(Len(str) - 1) As Byte

    Dim i As Integer
    For i = 1 To Len(str)
        b(i - 1) = Asc(Mid$(str, i, 1))
    Next

    stringToBytes = b
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
    CopyMemory ByteArray(0), ByVal VarPtr(intvalue), Len(intvalue)
    IntegerToByteArray = ByteArray

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
    CopyMemory ByteArray(0), ByVal VarPtr(lng), Len(lng)
    LongToByteArray = ByteArray

End Function

Function BytesToInteger(ByteArray() As Byte, FirstIndex As Long) As Integer
    Dim value As Long

    value = ByteArray(FirstIndex) + ByteArray(FirstIndex + 1) * 256&

    BytesToInteger = UnsignedToInteger(value)

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
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim lng256 As Double
    Dim lngReturn As Double

    lng256 = 1
    lngReturn = 0

    If EndRec < 1 Then
        EndRec = UBound(ByteArray)
    End If

    If StartRec > EndRec Or StartRec < 0 Then
        MessageBox _
                "Start record can not be greater then End record...!", _
                vbInformation
        BytesToNumEx = -1
        Exit Function
    End If

    lngReturn = lngReturn + (ByteArray(StartRec))
    For i = (StartRec + 1) To EndRec
        lng256 = lng256 * 256
        If i < EndRec Then
            lngReturn = lngReturn + (ByteArray(i) * lng256)
        Else
            ' if -ve

            If ByteArray(i) > 127 And UnSigned = False Then
                lngReturn = (lngReturn + ((ByteArray(i) - 256) _
                                          * lng256))
            Else
                lngReturn = lngReturn + (ByteArray(i) * lng256)
            End If
        End If
    Next i

    BytesToNumEx = lngReturn
ErrorHandler:
End Function
