VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Zlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Private Const m_def_CompressedSize = 0
Private Const m_def_OriginalSize = 0
'Property Variables:
Dim m_CompressedSize As Long
Dim m_OriginalSize As Long

'Declares
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Enum CZErrors
[Insufficient Buffer] = -5
End Enum



    '"Compress-Z-It" & Chr$(10) & Chr$(10) & "Data compression ActiveX component module." & Chr$(10) & Chr$(10) & "Custom control written and compiled by Benjamin Dowse. Portions written by other external 'zLib' compression software library authors." & Chr$(10) & Chr$(10) & "Special thanks and honor to the authors of the zLib DLL.", vbInformation + vbOKOnly, "DowseWare - Compress-Z-It ActiveX Control")



Public Function CompressData(TheData() As Byte) As Long

10    OriginalSize = UBound(TheData) + 1

      'Allocate memory for byte array
      Dim BufferSize As Long
      Dim TempBuffer() As Byte
      Dim result As Long

20    BufferSize = UBound(TheData) + 1
30    BufferSize = BufferSize + (BufferSize * 0.01) + 12
40    ReDim TempBuffer(BufferSize)

      'Compress byte array (data)
50    result = compress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

      'Truncate to compressed size
60    ReDim Preserve TheData(BufferSize - 1)
70    CopyMemory TheData(0), TempBuffer(0), BufferSize

      'Cleanup
80    Erase TempBuffer

      'Set properties if no error occurred
90    If result = 0 Then CompressedSize = UBound(TheData) + 1

      'Return error code (if any)
100   CompressData = result

End Function

'Public Function CompressString(TheString As String) As Long
'
'OriginalSize = Len(TheString)
'
''Allocate string space for the buffers
'Dim CmpSize As Long
'Dim TBuff As String
'orgSize = Len(TheString)
'TBuff = String(orgSize + (orgSize * 0.01) + 12, 0)
'CmpSize = Len(TBuff)
'
''Compress string (temporary string buffer) data
'ret = compress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))
'
''Set original value
'OriginalSize = Len(TheString)
'
''Crop the string and set it to the actual string.
'TheString = Left$(TBuff, CmpSize)
'
''Set compressed size of string.
'CompressedSize = CmpSize
'
''Cleanup
'TBuff = ""
'
''Return error code (if any)
'CompressString = ret
'
'End Function

Public Function DecompressData(TheData() As Byte, OrigSize As Long) As Long

      'Allocate memory for buffers
      Dim BufferSize As Long
      Dim TempBuffer() As Byte
      Dim result As Long

10    BufferSize = OrigSize
20    BufferSize = BufferSize + (BufferSize * 0.01) + 12
30    ReDim TempBuffer(BufferSize)

      'Decompress data
40    result = uncompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

      'Truncate buffer to compressed size
50    ReDim Preserve TheData(BufferSize - 1)
60    CopyMemory TheData(0), TempBuffer(0), BufferSize

      'Reset properties
70    If result = 0 Then
80    CompressedSize = 0
90    OriginalSize = 0
100   End If

      'Return error code (if any)
110   DecompressData = result

End Function

'Public Function DecompressString(TheString As String, OrigSize As Long) As Long
'
''Allocate string space
'Dim CmpSize As Long
'Dim TBuff As String
'TBuff = String(OriginalSize + (OriginalSize * 0.01) + 12, 0)
'CmpSize = Len(TBuff)
'
''Decompress
'result = uncompress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))
'
''Make string the size of the uncompressed string
'TheString = Left$(TBuff, CmpSize)
'
''Reset properties
'If result = 0 Then
'CompressedSize = 0
'OriginalSize = 0
'End If
'
''Return error code (if any)
'DecompressString = ret
'
'End Function

Public Property Get CompressedSize() As Long
10        CompressedSize = m_CompressedSize
End Property

Public Property Let CompressedSize(ByVal New_CompressedSize As Long)
10        m_CompressedSize = New_CompressedSize
End Property

Public Property Get OriginalSize() As Long
10        OriginalSize = m_OriginalSize
End Property

Public Property Let OriginalSize(ByVal New_OriginalSize As Long)
10        m_OriginalSize = New_OriginalSize
End Property

