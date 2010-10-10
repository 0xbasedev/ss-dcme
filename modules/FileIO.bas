Attribute VB_Name = "FileIO"
'Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
'' Some pages may also contain other copyrights by the author.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Distribution: You can freely use this code in your own
''               applications, but you may not reproduce
''               or publish this code on any web site,
''               online service, or distribute as source
''               on any media without express permission.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Const GENERIC_READ As Long = &H80000000
'Private Const INVALID_HANDLE_VALUE As Long = -1
'Private Const OPEN_EXISTING As Long = 3
'Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
'Private Const MAX_PATH As Long = 260
'
''Enum containing values representing
''the status of the file
'Enum IsFileResults
'   FILE_IN_USE = -1  'True
'   FILE_FREE = 0     'False
'   FILE_DOESNT_EXIST = -999 'arbitrary number, other than 0 or -1
'End Enum
'
'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type WIN32_FIND_DATA
'   dwFileAttributes As Long
'   ftCreationTime As FILETIME
'   ftLastAccessTime As FILETIME
'   ftLastWriteTime As FILETIME
'   nFileSizeHigh As Long
'   nFileSizeLow As Long
'   dwReserved0 As Long
'   dwReserved1 As Long
'   cFileName As String * MAX_PATH
'   cAlternate As String * 14
'End Type
'
'Private Declare Function CreateFile Lib "kernel32" _
'   Alias "CreateFileA" _
'  (ByVal lpFileName As String, _
'   ByVal dwDesiredAccess As Long, _
'   ByVal dwShareMode As Long, _
'   ByVal lpSecurityAttributes As Long, _
'   ByVal dwCreationDisposition As Long, _
'   ByVal dwFlagsAndAttributes As Long, _
'   ByVal hTemplateFile As Long) As Long
'
'Private Declare Function CloseHandle Lib "kernel32" _
'  (ByVal hFile As Long) As Long
'
'Private Declare Function FindFirstFile Lib "kernel32" _
'   Alias "FindFirstFileA" _
'  (ByVal lpFileName As String, _
'   lpFindFileData As WIN32_FIND_DATA) As Long
'
'Private Declare Function FindClose Lib "kernel32" _
'  (ByVal hFindFile As Long) As Long
'
'
'
'Function IsFileInUse(sFile As String) As IsFileResults
'
'   Dim hFile As Long
'
'   If FileExists(sFile) Then
'
'     'note that FILE_ATTRIBUTE_NORMAL (&H80) has
'     'a different value than VB's constant vbNormal (0)!
'      hFile = CreateFile(sFile, _
'                         GENERIC_READ, _
'                         0, 0, _
'                         OPEN_EXISTING, _
'                         FILE_ATTRIBUTE_NORMAL, 0&)
'
'     'this will evaluate to either
'     '-1 (FILE_IN_USE) or 0 (FILE_FREE)
'      IsFileInUse = hFile = INVALID_HANDLE_VALUE
'
'      CloseHandle hFile
'
'   Else
'
'     'the value of FILE_DOESNT_EXIST in the Enum
'     'is arbitrary, as long as it's not 0 or -1
'      IsFileInUse = FILE_DOESNT_EXIST
'
'   End If
'
'End Function
'
'
'Function FileExists(sSource As String) As Boolean
'
'   Dim WFD As WIN32_FIND_DATA
'   Dim hFile As Long
'
'   hFile = FindFirstFile(sSource, WFD)
'   FileExists = hFile <> INVALID_HANDLE_VALUE
'
'   Call FindClose(hFile)
'
'End Function
'
'
