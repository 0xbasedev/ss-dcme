Attribute VB_Name = "userInfo"
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish

Public Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
 
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

'50         Select Case .dwPlatformId
'
'            Case 1
'
'60              Select Case .dwMinorVersion
'                    Case 0
'70                      getWindowsVersion = os_win95
'80                  Case 10
'90                      getWindowsVersion = os_win98
'100                 Case 90
'110                     getWindowsVersion = os_winME
'120             End Select
'
'130         Case 2
'140             Select Case .dwMajorVersion
'                    Case 3
'150                     getWindowsVersion = os_winNT35
'160                 Case 4
'170                     getWindowsVersion = os_winNT4
'180                 Case 5
'190                     If .dwMinorVersion = 0 Then
'200                         getWindowsVersion = os_win2000
'210                     Else
'220                         getWindowsVersion = os_winxp
'230                     End If
'240                   Case 6
'250                       getWindowsVersion = os_winvista
'260              End Select

Enum WindowsVersionEnum
    os_win95 = 10
    os_win98 = 11
    os_winME = 19
    os_winNT35 = 30
    os_winNT4 = 40
    os_win2000 = 50
    os_winxp = 51
    os_winvista = 60
    os_win7 = 70
    os_unknown = 0
End Enum

Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" _
(ByVal lpBuffer As String, _
nSize As Long) As Long

Private Type USER_INFO_2
    usri2_name As Long
    usri2_password As Long    ' Null, only settable
    usri2_password_age As Long
    usri2_priv As Long
    usri2_home_dir As Long
    usri2_comment As Long
    usri2_flags As Long
    usri2_script_path As Long
    usri2_auth_flags As Long
    usri2_full_name As Long
    usri2_usr_comment As Long
    usri2_parms As Long
    usri2_workstations As Long
    usri2_last_logon As Long
    usri2_last_logoff As Long
    usri2_acct_expires As Long
    usri2_max_storage As Long
    usri2_units_per_week As Long
    usri2_logon_hours As Long
    usri2_bad_pw_count As Long
    usri2_num_logons As Long
    usri2_logon_server As Long
    usri2_country_code As Long
    usri2_code_page As Long
End Type

Private Declare Function apiNetGetDCName _
                          Lib "netapi32.dll" Alias "NetGetDCName" _
                              (ByVal servername As Long, _
                               ByVal DomainName As Long, _
                               bufptr As Long) As Long

' function frees the memory that the NetApiBufferAllocate
' function allocates.
Private Declare Function apiNetAPIBufferFree _
                          Lib "netapi32.dll" Alias "NetApiBufferFree" _
                              (ByVal Buffer As Long) _
                              As Long

' Retrieves the length of the specified wide string.
Private Declare Function apilstrlenW _
                          Lib "kernel32" Alias "lstrlenW" _
                              (ByVal lpString As Long) _
                              As Long

Private Declare Function apiNetUserGetInfo _
                          Lib "netapi32.dll" Alias "NetUserGetInfo" _
                              (servername As Any, _
                               username As Any, _
                               ByVal Level As Long, _
                               bufptr As Long) As Long

' moves memory either forward or backward, aligned or unaligned,
' in 4-byte blocks, followed by any remaining bytes
Private Declare Sub sapiCopyMem _
                     Lib "kernel32" Alias "RtlMoveMemory" _
                         (Destination As Any, _
                          Source As Any, _
                          ByVal Length As Long)

Private Declare Function apiGetUserName Lib _
                                        "advapi32.dll" Alias "GetUserNameA" _
                                        (ByVal lpBuffer As String, _
                                         nSize As Long) _
                                         As Long

Private Const MAXCOMMENTSZ = 256
Private Const NERR_SUCCESS = 0
Private Const ERROR_MORE_DATA = 234&
Private Const MAX_CHUNK = 25
Private Const ERROR_SUCCESS = 0&



Private Enum EXTENDED_NAME_FORMAT
NameUnknown = 0
NameFullyQualifiedDN = 1
NameSamCompatible = 2
NameDisplay = 3
NameUniqueId = 6
NameCanonical = 7
NameUserPrincipal = 8
NameCanonicalEx = 9
NameServicePrincipal = 10
End Enum
Private Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long

Public currentWindowsVersion As WindowsVersionEnum

Private Function fGetFullNameOfLoggedUser(Optional strUserName As String) As String
'
' Returns the full name for a given UserID
'   NT/2000 only
' Omitting the strUserName argument will try and
' retrieve the full name for the currently logged on user
'
    On Error GoTo ErrHandler
    Dim pBuf As Long
    Dim dwRec As Long
    Dim pTmp As USER_INFO_2
    Dim abytPDCName() As Byte
    Dim abytUserName() As Byte
    Dim lngRet As Long
    Dim i As Long

    ' Unicode
    abytPDCName = fGetDCName() & vbNullChar
    If (Len(strUserName) = 0) Then strUserName = fGetUserName()
    abytUserName = strUserName & vbNullChar

    ' Level 2
    lngRet = apiNetUserGetInfo( _
             abytPDCName(0), _
             abytUserName(0), _
             2, _
             pBuf)
    If (lngRet = ERROR_SUCCESS) Then
        Call sapiCopyMem(pTmp, ByVal pBuf, Len(pTmp))
        fGetFullNameOfLoggedUser = fStrFromPtrW(pTmp.usri2_full_name)
    End If

    Call apiNetAPIBufferFree(pBuf)
ExitHere:
    Exit Function
ErrHandler:
    fGetFullNameOfLoggedUser = vbNullString
    Resume ExitHere
End Function

Private Function fGetUserName() As String
' Returns the network login name
    Dim lngLen As Long, lngRet As Long
    Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngRet = apiGetUserName(strUserName, lngLen)
    If lngRet Then
        fGetUserName = Left$(strUserName, lngLen - 1)
    End If
End Function

Private Function fGetDCName() As String
    Dim pTmp As Long
    Dim lngRet As Long
    ' Dim abytBuf() As Byte

    lngRet = apiNetGetDCName(0, 0, pTmp)
    If lngRet = NERR_SUCCESS Then
        fGetDCName = fStrFromPtrW(pTmp)
    End If
    Call apiNetAPIBufferFree(pTmp)
End Function

Private Function fStrFromPtrW(pBuf As Long) As String
    Dim lngLen As Long
    Dim abytBuf() As Byte

    ' Get the length of the string at the memory location
    lngLen = apilstrlenW(pBuf) * 2
    ' if it's not a ZLS
    If lngLen Then
        ReDim abytBuf(lngLen)
        ' then copy the memory contents
        ' into a temp buffer
        Call sapiCopyMem( _
             abytBuf(0), _
             ByVal pBuf, _
             lngLen)
        ' return the buffer
        fStrFromPtrW = abytBuf
    End If
End Function

Function fGetUserNameEx() As String
    Dim sBuffer As String, ret As Long
    sBuffer = String(256, 0)
    ret = Len(sBuffer)
    
    If GetUserNameEx(NameSamCompatible, sBuffer, ret) <> 0 Then
        fGetUserNameEx = Left$(sBuffer, ret)
    Else
        fGetUserNameEx = ""
    End If
End Function

Private Function fGetComputerName() As String

    'return the name of the computer
    Dim tmp As String
    
    tmp = Space$(MAX_COMPUTERNAME + 1)
    
    If GetComputerName(tmp, Len(tmp)) <> 0 Then
        fGetComputerName = TrimNull(tmp)
    Else
        fGetComputerName = ""
    End If

End Function


Private Function TrimNull(item As String)

Dim pos As Integer

pos = InStr(item, Chr$(0))

If pos Then
TrimNull = Left$(item, pos - 1)
Else: TrimNull = item
End If

End Function

Function GetUserName() As String
    Dim ret As String
    ret = fGetDCName

    If ret <> "" Then
        GetUserName = ret
    Else
        GetUserName = fGetUserName
    End If
End Function

Function getWindowsVersion() As WindowsVersionEnum
    Dim osinfo As OSVERSIONINFO
    
    Dim retvalue As Integer

     osinfo.dwOSVersionInfoSize = 148
     osinfo.szCSDVersion = Space$(128)
     retvalue = GetVersionExA(osinfo)

     With osinfo
     Select Case .dwPlatformId

      Case 1
      
          Select Case .dwMinorVersion
              Case 0
                  getWindowsVersion = os_win95
              Case 10
                  getWindowsVersion = os_win98
              Case 90
                  getWindowsVersion = os_winME
              Case Else
                  getWindowsVersion = os_unknown
          End Select

      Case 2
          Select Case .dwMajorVersion
              Case 3
                  getWindowsVersion = os_winNT35
              Case 4
                  getWindowsVersion = os_winNT4
              Case 5
                  If .dwMinorVersion = 0 Then
                      getWindowsVersion = os_win2000
                  Else
                      getWindowsVersion = os_winxp
                  End If
              Case 6
                  getWindowsVersion = os_winvista
              Case Else
                  getWindowsVersion = os_win7
           End Select
      Case Else
          getWindowsVersion = os_unknown
        End Select
    End With
    
End Function









