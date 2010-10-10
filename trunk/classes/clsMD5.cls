VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is a very minor adaption of the work of Mark van Cuijk
' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=41067&lngWId=1

Private lngTrack As Long
Private arrLongConversion(4) As Long
Private arrSplit64(63) As Byte

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

Private Function MD5Round(strRound As String, a As Long, b As Long, c As Long, d As Long, X As Long, S As Long, ac As Long) As Long
10        Select Case strRound
              Case Is = "FF"
20                a = MD5LongAdd4(a, (b And c) Or (Not (b) And d), X, ac)
30                a = MD5Rotate(a, S)
40                a = MD5LongAdd(a, b)
50            Case Is = "GG"
60                a = MD5LongAdd4(a, (b And d) Or (c And Not (d)), X, ac)
70                a = MD5Rotate(a, S)
80                a = MD5LongAdd(a, b)
90            Case Is = "HH"
100               a = MD5LongAdd4(a, b Xor c Xor d, X, ac)
110               a = MD5Rotate(a, S)
120               a = MD5LongAdd(a, b)
130           Case Is = "II"
140               a = MD5LongAdd4(a, c Xor (b Or Not (d)), X, ac)
150               a = MD5Rotate(a, S)
160               a = MD5LongAdd(a, b)
170       End Select
End Function

Private Function MD5Rotate(lngValue As Long, lngBits As Long) As Long
      Dim lngSign As Long
      Dim lngI As Long
10        lngBits = (lngBits Mod 32)
          
20        If lngBits = 0 Then MD5Rotate = lngValue: Exit Function
          
30        For lngI = 1 To lngBits
40            lngSign = lngValue And &HC0000000
50            lngValue = (lngValue And &H3FFFFFFF) * 2
60            lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
70        Next
          
80        MD5Rotate = lngValue
End Function

Private Function TRID() As String

          Dim sngNum As Single, lngnum As Long
          Dim strResult As String
         
10        sngNum = Rnd(2147483648#)
20        strResult = CStr(sngNum)
          
30        strResult = replace(strResult, "0.", "")
40        strResult = replace(strResult, ".", "")
50        strResult = replace(strResult, "E-", "")
          
60        TRID = strResult

End Function

Private Function MD564Split(lngLength As Long, bytBuffer() As Byte) As String

          Dim lngBytesTotal As Long, lngBytesToAdd As Long
          Dim intLoop As Integer, intLoop2 As Integer, lngTrace As Long
          Dim intInnerLoop As Integer, intLoop3 As Integer
          
10        lngBytesTotal = lngTrack Mod 64
20        lngBytesToAdd = 64 - lngBytesTotal
30        lngTrack = (lngTrack + lngLength)
          
40        If lngLength >= lngBytesToAdd Then
50            For intLoop = 0 To lngBytesToAdd - 1
60                arrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
70            Next intLoop
              
80            MD5Conversion arrSplit64
              
90            lngTrace = (lngLength) Mod 64

100           For intLoop2 = lngBytesToAdd To (lngLength - intLoop - lngTrace) Mod 32767 Step 64
110               For intInnerLoop = 0 To 63
120                   arrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
130               Next intInnerLoop
                  
140               MD5Conversion arrSplit64
              
150           Next intLoop2
              
160           lngBytesTotal = 0
170       Else
          
180         intLoop2 = 0
          
190       End If
          
200       For intLoop3 = 0 To (lngLength - intLoop2 - 1) Mod 32767
              
210           arrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)
          
220       Next intLoop3
           
End Function

Private Function MD5StringArray(strInput As String) As Byte()
          
          Dim intLoop As Integer
          Dim bytBuffer() As Byte
10        ReDim bytBuffer(Len(strInput))
          
20        For intLoop = 0 To Len(strInput) - 1
30            bytBuffer(intLoop) = Asc(Mid(strInput, intLoop + 1, 1))
40        Next intLoop
          
50        MD5StringArray = bytBuffer
          
End Function

Private Sub MD5Conversion(bytBuffer() As Byte)

          Dim X(16) As Long, a As Long
          Dim b As Long, c As Long
          Dim d As Long
          
10        a = arrLongConversion(1)
20        b = arrLongConversion(2)
30        c = arrLongConversion(3)
40        d = arrLongConversion(4)
          
50        MD5Decode 64, X, bytBuffer
          
60        MD5Round "FF", a, b, c, d, X(0), S11, -680876936
70        MD5Round "FF", d, a, b, c, X(1), S12, -389564586
80        MD5Round "FF", c, d, a, b, X(2), S13, 606105819
90        MD5Round "FF", b, c, d, a, X(3), S14, -1044525330
100       MD5Round "FF", a, b, c, d, X(4), S11, -176418897
110       MD5Round "FF", d, a, b, c, X(5), S12, 1200080426
120       MD5Round "FF", c, d, a, b, X(6), S13, -1473231341
130       MD5Round "FF", b, c, d, a, X(7), S14, -45705983
140       MD5Round "FF", a, b, c, d, X(8), S11, 1770035416
150       MD5Round "FF", d, a, b, c, X(9), S12, -1958414417
160       MD5Round "FF", c, d, a, b, X(10), S13, -42063
170       MD5Round "FF", b, c, d, a, X(11), S14, -1990404162
180       MD5Round "FF", a, b, c, d, X(12), S11, 1804603682
190       MD5Round "FF", d, a, b, c, X(13), S12, -40341101
200       MD5Round "FF", c, d, a, b, X(14), S13, -1502002290
210       MD5Round "FF", b, c, d, a, X(15), S14, 1236535329

220       MD5Round "GG", a, b, c, d, X(1), S21, -165796510
230       MD5Round "GG", d, a, b, c, X(6), S22, -1069501632
240       MD5Round "GG", c, d, a, b, X(11), S23, 643717713
250       MD5Round "GG", b, c, d, a, X(0), S24, -373897302
260       MD5Round "GG", a, b, c, d, X(5), S21, -701558691
270       MD5Round "GG", d, a, b, c, X(10), S22, 38016083
280       MD5Round "GG", c, d, a, b, X(15), S23, -660478335
290       MD5Round "GG", b, c, d, a, X(4), S24, -405537848
300       MD5Round "GG", a, b, c, d, X(9), S21, 568446438
310       MD5Round "GG", d, a, b, c, X(14), S22, -1019803690
320       MD5Round "GG", c, d, a, b, X(3), S23, -187363961
330       MD5Round "GG", b, c, d, a, X(8), S24, 1163531501
340       MD5Round "GG", a, b, c, d, X(13), S21, -1444681467
350       MD5Round "GG", d, a, b, c, X(2), S22, -51403784
360       MD5Round "GG", c, d, a, b, X(7), S23, 1735328473
370       MD5Round "GG", b, c, d, a, X(12), S24, -1926607734
        
380       MD5Round "HH", a, b, c, d, X(5), S31, -378558
390       MD5Round "HH", d, a, b, c, X(8), S32, -2022574463
400       MD5Round "HH", c, d, a, b, X(11), S33, 1839030562
410       MD5Round "HH", b, c, d, a, X(14), S34, -35309556
420       MD5Round "HH", a, b, c, d, X(1), S31, -1530992060
430       MD5Round "HH", d, a, b, c, X(4), S32, 1272893353
440       MD5Round "HH", c, d, a, b, X(7), S33, -155497632
450       MD5Round "HH", b, c, d, a, X(10), S34, -1094730640
460       MD5Round "HH", a, b, c, d, X(13), S31, 681279174
470       MD5Round "HH", d, a, b, c, X(0), S32, -358537222
480       MD5Round "HH", c, d, a, b, X(3), S33, -722521979
490       MD5Round "HH", b, c, d, a, X(6), S34, 76029189
500       MD5Round "HH", a, b, c, d, X(9), S31, -640364487
510       MD5Round "HH", d, a, b, c, X(12), S32, -421815835
520       MD5Round "HH", c, d, a, b, X(15), S33, 530742520
530       MD5Round "HH", b, c, d, a, X(2), S34, -995338651
       
540       MD5Round "II", a, b, c, d, X(0), S41, -198630844
550       MD5Round "II", d, a, b, c, X(7), S42, 1126891415
560       MD5Round "II", c, d, a, b, X(14), S43, -1416354905
570       MD5Round "II", b, c, d, a, X(5), S44, -57434055
580       MD5Round "II", a, b, c, d, X(12), S41, 1700485571
590       MD5Round "II", d, a, b, c, X(3), S42, -1894986606
600       MD5Round "II", c, d, a, b, X(10), S43, -1051523
610       MD5Round "II", b, c, d, a, X(1), S44, -2054922799
620       MD5Round "II", a, b, c, d, X(8), S41, 1873313359
630       MD5Round "II", d, a, b, c, X(15), S42, -30611744
640       MD5Round "II", c, d, a, b, X(6), S43, -1560198380
650       MD5Round "II", b, c, d, a, X(13), S44, 1309151649
660       MD5Round "II", a, b, c, d, X(4), S41, -145523070
670       MD5Round "II", d, a, b, c, X(11), S42, -1120210379
680       MD5Round "II", c, d, a, b, X(2), S43, 718787259
690       MD5Round "II", b, c, d, a, X(9), S44, -343485551
          
700       arrLongConversion(1) = MD5LongAdd(arrLongConversion(1), a)
710       arrLongConversion(2) = MD5LongAdd(arrLongConversion(2), b)
720       arrLongConversion(3) = MD5LongAdd(arrLongConversion(3), c)
730       arrLongConversion(4) = MD5LongAdd(arrLongConversion(4), d)
          
End Sub

Private Function MD5LongAdd(lngVal1 As Long, lngVal2 As Long) As Long
          
          Dim lngHighWord As Long
          Dim lngLowWord As Long
          Dim lngOverflow As Long

10        lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
20        lngOverflow = lngLowWord \ 65536
30        lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
          
40        MD5LongAdd = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Function MD5LongAdd4(lngVal1 As Long, lngVal2 As Long, lngVal3 As Long, lngVal4 As Long) As Long
          
          Dim lngHighWord As Long
          Dim lngLowWord As Long
          Dim lngOverflow As Long

10        lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
20        lngOverflow = lngLowWord \ 65536
30        lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
40        MD5LongAdd4 = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Sub MD5Decode(intLength As Integer, lngOutBuffer() As Long, bytInBuffer() As Byte)
          
          Dim intDblIndex As Integer
          Dim intByteIndex As Integer
          Dim dblSum As Double
          
10        intDblIndex = 0
          
20        For intByteIndex = 0 To intLength - 1 Step 4
              
30            dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256# + bytInBuffer(intByteIndex + 2) * 65536# + bytInBuffer(intByteIndex + 3) * 16777216#
40            lngOutBuffer(intDblIndex) = MD5LongConversion(dblSum)
50            intDblIndex = (intDblIndex + 1)
          
60        Next intByteIndex

End Sub

Private Function MD5LongConversion(dblValue As Double) As Long
          
10        If dblValue < 0 Or dblValue >= OFFSET_4 Then Error 6
              
20        If dblValue <= MAXINT_4 Then
30            MD5LongConversion = dblValue
40        Else
50            MD5LongConversion = dblValue - OFFSET_4
60        End If
              
End Function

Private Sub MD5Finish()
      Dim dblBits As Double
      Dim arrPadding(72) As Byte
      Dim lngBytesBuffered As Long
          
10        arrPadding(0) = &H80
20        dblBits = lngTrack * 8
          
30        lngBytesBuffered = lngTrack Mod 64
          
40        If lngBytesBuffered <= 56 Then
50            MD564Split (56 - lngBytesBuffered), arrPadding
60        Else
70            MD564Split (120 - lngTrack), arrPadding
80        End If
          
          
90        arrPadding(0) = MD5LongConversion(dblBits) And &HFF&
100       arrPadding(1) = MD5LongConversion(dblBits) \ 256 And &HFF&
110       arrPadding(2) = MD5LongConversion(dblBits) \ 65536 And &HFF&
120       arrPadding(3) = MD5LongConversion(dblBits) \ 16777216 And &HFF&
130       arrPadding(4) = 0
140       arrPadding(5) = 0
150       arrPadding(6) = 0
160       arrPadding(7) = 0
          
170       MD564Split 8, arrPadding
End Sub

Private Function MD5StringChange(lngnum As Long) As String
      Dim bytA As Byte
      Dim bytB As Byte
      Dim bytC As Byte
      Dim bytD As Byte
10         bytA = lngnum And &HFF&
20         If bytA < 16 Then
30             MD5StringChange = "0" & Hex(bytA)
40         Else
50             MD5StringChange = Hex(bytA)
60         End If
                  
70         bytB = (lngnum And &HFF00&) \ 256
80         If bytB < 16 Then
90             MD5StringChange = MD5StringChange & "0" & Hex(bytB)
100        Else
110            MD5StringChange = MD5StringChange & Hex(bytB)
120        End If
           
130        bytC = (lngnum And &HFF0000) \ 65536
140        If bytC < 16 Then
150            MD5StringChange = MD5StringChange & "0" & Hex(bytC)
160        Else
170            MD5StringChange = MD5StringChange & Hex(bytC)
180        End If
          
190        If lngnum < 0 Then
200            bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
210        Else
220            bytD = (lngnum And &HFF000000) \ 16777216
230        End If
           
240        If bytD < 16 Then
250            MD5StringChange = MD5StringChange & "0" & Hex(bytD)
260        Else
270            MD5StringChange = MD5StringChange & Hex(bytD)
280        End If
End Function

Function GetMD5Value(hash() As Long) As String
10        If UBound(hash) = 3 Then
20            GetMD5Value = LCase(MD5StringChange(hash(0)) & MD5StringChange(hash(1)) & MD5StringChange(hash(2)) & MD5StringChange(hash(3)))
30        ElseIf UBound(hash) = 4 Then
40            GetMD5Value = LCase(MD5StringChange(hash(1)) & MD5StringChange(hash(2)) & MD5StringChange(hash(3)) & MD5StringChange(hash(4)))
50        Else
60            GetMD5Value = "ERROR: Ubound=" & UBound(hash)
70        End If
End Function

Function MD5Value() As String
10        MD5Value = GetMD5Value(arrLongConversion)
End Function

Public Function CalculateMD5(filename As String) As Long()
          Dim bytBuffer() As Byte
10        If Not FileExists(filename) Then
20            ReDim bytBuffer(0)
30            bytBuffer(0) = 0
40        Else
50            ReDim bytBuffer(FileLen(filename) - 1)
              
              Dim f As Integer
60            f = FreeFile
70            Open filename For Binary As #f
80            Get #f, , bytBuffer
90            Close #f
              
100       End If
          
110       MD5Start
120       MD564Split UBound(bytBuffer) + 1, bytBuffer
130       MD5Finish
          
          Dim retlng() As Long
140       ReDim retlng(3)
          
          Dim i As Integer
150       For i = 0 To 3
160           retlng(i) = arrLongConversion(i + 1)
170       Next
          
180       CalculateMD5 = retlng
          
End Function

Public Function CompareWith(testMD5() As Long) As Boolean
10        If UBound(arrLongConversion) = 3 And UBound(testMD5) = 3 Then
              Dim i As Integer
              
20            For i = 0 To 3
30                If arrLongConversion(i) <> testMD5(i) Then
40                    CompareWith = False
50                    Exit Function
60                End If
70            Next
80        Else
90            CompareWith = False
100       End If
110       CompareWith = True
End Function

Private Sub MD5Start()
10        lngTrack = 0
20        arrLongConversion(1) = MD5LongConversion(1732584193#)
30        arrLongConversion(2) = MD5LongConversion(4023233417#)
40        arrLongConversion(3) = MD5LongConversion(2562383102#)
50        arrLongConversion(4) = MD5LongConversion(271733878#)
End Sub

