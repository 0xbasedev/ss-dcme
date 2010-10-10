Attribute VB_Name = "mdlRegistry"
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long



Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
                              "RegCreateKeyA" (ByVal hKey As Long, _
                                               ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias _
                             "RegSetValueA" (ByVal hKey As Long, _
                                             ByVal lpSubKey As String, ByVal dwType As Long, _
                                             ByVal lpData As String, ByVal cbData As Long) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, Source As Any, ByVal numBytes As Long)

Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))


Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234


' Return codes from Registration functions.
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1&
Public Const ERROR_BADKEY = 2&
Public Const ERROR_CANTOPEN = 3&
Public Const ERROR_CANTREAD = 4&
Public Const ERROR_CANTWRITE = 5&
Public Const ERROR_OUTOFMEMORY = 6&
Public Const ERROR_INVALID_PARAMETER = 7&
Public Const ERROR_ACCESS_DENIED = 8&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 260&
Public Const REG_SZ = 1


Sub AssignExt()
    Dim sKeyName As String   'Holds Key Name in registry.
    Dim sKeyValue As String  'Holds Key Value in registry.
    Dim ret&           'Holds error status if any from API calls.
    Dim lphKey&        'Holds created key handle from RegCreateKey.
    
'This creates a Root entry called "MyApp".
    sKeyName = "DCME"
    sKeyValue = "LVL File"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

    'This creates a Root entry called .XXX associated with "MyApp".
    'You can replace ".XXX" with your wanted extension
    sKeyName = ".LVL"
    'replace all "MyApp" below with your application name
    sKeyValue = "DCME"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

    'This sets the command line for "MyApp".
    sKeyName = "DCME"
    'replace c:\mydir\my.exe with your exe file. In this example,
    'All the .XXX files will be opened with the file c:\mydir\my.exe
    sKeyValue = """" & GetApplicationFullPath & """ ""%1"""
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, _
                        sKeyValue, MAX_PATH)
    '*******************************************************
    'That's All!
    'Now to test this program, after you associate the .XXX files with
    'c:\mydir\my.exe (In this example), start a new project, and
    'copy the following 3 lines to your form (uncomment the lines)

    'Private Sub Form_Load()
    '    messagebox Command
    'End Sub

    'compile the program to my.exe file, and put it in c:\mydir directory.
    'Now go back to Windows and change the name of one of your files
    'To   Test.xxx   and  double click on it. It will be opened with the program
    'c:\mydir\my.exe and it will pop up message box: "c:\Test.xxx"
    '***********************************************************

    MessageBox "Done!", vbInformation

End Sub

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Optional defaultValue As Variant) As Variant
    Dim Handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim resBinary() As Byte
    Dim Length As Long
    Dim retVal As Long
    Dim valueType As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(defaultValue), Empty, defaultValue)
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, Handle) Then
        Exit Function
    End If
    
    ' prepare a 1K receiving resBinary
    Length = 1024
    ReDim resBinary(0 To Length - 1) As Byte
    
    ' read the registry key
    retVal = RegQueryValueEx(Handle, ValueName, 0, valueType, resBinary(0), _
        Length)
    ' if resBinary was too small, try again
    If retVal = ERROR_MORE_DATA Then
        ' enlarge the resBinary, and read the value again
        ReDim resBinary(0 To Length - 1) As Byte
        retVal = RegQueryValueEx(Handle, ValueName, 0, valueType, resBinary(0), _
            Length)
    End If
    
    ' return a value corresponding to the value type
    Select Case valueType
        Case REG_DWORD
            CopyMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case REG_SZ, REG_EXPAND_SZ
            ' copy everything but the trailing null char
            resString = Space$(Length - 1)
            CopyMemory ByVal resString, resBinary(0), Length - 1
            GetRegistryValue = resString
        Case REG_BINARY
            ' resize the result resBinary
            If Length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To Length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case REG_MULTI_SZ
            ' copy everything but the 2 trailing null chars
            resString = Space$(Length - 2)
            CopyMemory ByVal resString, resBinary(0), Length - 2
            GetRegistryValue = resString
        Case Else
            RegCloseKey Handle
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' close the registry key
    RegCloseKey Handle
End Function







Function FindAssociatedProgram(ByVal extension As _
    String) As String
    
    Dim temp_title As String
    Dim temp_path As String
    Dim fnum As Integer
    Dim result As String
    Dim pos As Integer

    ' Get a temporary file name with this extension.
    GetTempFile extension, temp_path, temp_title

    ' Make the file.
    fnum = FreeFile
    Open temp_path & temp_title For Output As fnum
    Close fnum

    ' Get the associated executable.
    result = Space$(1024)
    FindExecutable temp_title, temp_path, result
    pos = InStr(result, Chr$(0))
    FindAssociatedProgram = Left$(result, pos - 1)

    ' Delete the temporary file.
    Kill temp_path & temp_title
End Function

' Return a temporary file name.
Private Sub GetTempFile(ByVal extension As String, ByRef _
    temp_path As String, ByRef temp_title As String)
    
    Dim i As Integer

    If Left$(extension, 1) <> "." Then extension = "." & _
        extension

    temp_path = Environ("TEMP")
    If Right$(temp_path, 1) <> "\" Then temp_path = _
        temp_path & "\"

    i = 0
    Do
        temp_title = "tmp" & Format$(i) & extension
        If Len(Dir$(temp_path & temp_title)) = 0 Then Exit _
            Do
        i = i + 1
    Loop
End Sub





Function IsLVLAssociatedToDCME() As Boolean
    IsLVLAssociatedToDCME = (LCase(FindAssociatedProgram("lvl")) = LCase(GetApplicationFullPath))
End Function

Function GetSystemImageEditor() As String
    'Returns the full path of the system's default image editor
    'If bitmaps are opened with 'Image Preview', this will return the path of mspaint
    
    Dim ret As String
    ret = FindAssociatedProgram("bmp")
    
    If GetExtension(ret) <> "exe" Then
        'Use paint
        GetSystemImageEditor = SysDir & "\mspaint.exe"
    Else
        GetSystemImageEditor = ret
    End If
End Function
