VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheckModifs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************************************************
'Author: Yang Kok Wah
'Date: 22 Aug 2003
'Email:kwyang@starhub.net.sg
'The author claim no responsibility for any problems or liabilities
'arising from the use of 'the following code:
'************************************************************************
Option Explicit

Private Const TIME_OUT = &H102
Private Const FILE_SHARE_DELETE = &H4
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ALL_ACCESS = &H1FF

Private Const FILE_LIST_DIRECTORY = &H1
Private Const OPEN_EXISTING = &H3
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Enum FILE_NOTIFY_CHANGE
    FILE_NOTIFY_CHANGE_FILE_NAME = &H1
    FILE_NOTIFY_CHANGE_DIR_NAME = &H2
    FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
    FILE_NOTIFY_CHANGE_SIZE = &H8
    FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
    FILE_NOTIFY_CHANGE_LAST_ACCESS = &H20
    FILE_NOTIFY_CHANGE_CREATION = &H40
    FILE_NOTIFY_CHANGE_SECURITY = &H100
End Enum

Enum FILE_ACTION
    FILE_ACTION_ADDED = &H1
    FILE_ACTION_REMOVED = &H2
    FILE_ACTION_MODIFIED = &H3
    FILE_ACTION_RENAMED_OLD_NAME = &H4
    FILE_ACTION_RENAMED_NEW_NAME = &H5
End Enum

'comment off : not used
'Private Type SECURITY_ATTRIBUTES
' nLength As Long
' lpSecurityDescriptor As Long
' bInheritHandle As Long
'End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type FILE_NOTIFY_INFORMATION
    dwNextEntryOffset As Long
    dwAction As FILE_ACTION
    dwFileNameLength As Long
    wcFileName(1024 - 1) As Byte 'buffer of 1024 bytes
End Type

Private Declare Function ResetEvent Lib "kernel32" _
    (ByVal hEvent As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Declare Function CreateEvent Lib "kernel32" _
           Alias "CreateEventA" _
    (ByVal lpEventAttributes As Long, _
                ByVal bManualReset As Long, _
                ByVal bInitialState As Long, _
                ByVal lpName As String) As Long

Private Declare Function GetOverlappedResult Lib "kernel32" _
               (ByVal hFile As Long, lpOverlapped As OVERLAPPED, _
               lpNumberOfBytesTransferred As Long, _
               ByVal bWait As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long


Private Declare Function CreateFile Lib "kernel32" _
    Alias "CreateFileA" _
 (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
  ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
  ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, _
  ByVal hTemplateFile As Long) As Long


Private Declare Function ReadDirectoryChangesW Lib "kernel32.dll" _
    (ByVal hDirectory As Long, _
     ByVal lpBuffer As Long, _
     ByVal nBufferLength As Long, _
     ByVal bWatchSubtree As Boolean, _
    ByVal dwNotifyFilter As FILE_NOTIFY_CHANGE, _
    lpBytesReturned As Long, _
    ByVal lpOverlapped As Long, _
    ByVal lpCompletionRoutine As Long) As Long



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)


'Flag to stop mointoring
Dim fstop As Boolean



Dim hDir As Long 'Directory Handler
Dim hEvent As Long 'Event Handler

'Create an OverLapped structure for Async mode
Dim oLap As OVERLAPPED

'Buffer for ReadDirectoryChangesW
Dim buf(0 To 1024 * 5 - 1) As Byte

'structure to retreive file information
Dim dirBuf As FILE_NOTIFY_INFORMATION

Dim nUsed As Long 'used only for sync mode
    
Event FileModified(filename As String)
Event FileDeleted(filename As String)

Event FileAction(filename As String, FA As FILE_ACTION)



Sub StopChecking()
    'close all the handles
    CloseHandle hEvent
    CloseHandle hDir
End Sub

Sub InitCheck(filename As String)
    'Create the handle to mointor "c:\test"
    hDir = CreateFile(filename, _
                      0&, _
                      FILE_SHARE_READ Or _
                      FILE_SHARE_DELETE Or FILE_SHARE_WRITE, _
                      0&, _
                      OPEN_EXISTING, _
                      FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, _
                      0&)

    'Documenting CreateEvent
    'Create an Event for Async mode in ReadDirectoryChangesW
    'hEvent = CreateEvent( _
    ' 0&, _ 'No Security Attribute
    ' True, _ 'Manual Reset
    ' True, _ 'Initial State is Set
    ' "DirEvent") 'Unique name of Event
    
    hEvent = CreateEvent( _
            0&, _
            True, _
            True, _
            "DirEvent")
              
    'Assign the event created to the Overlapped structure
    With oLap
     .hEvent = hEvent
    End With

    
    'Documenting ReadDirectoryChangesW call
    'ReadDirectoryChangesW in Asyn mode
    'Call ReadDirectoryChangesW(hDir, _ ' Directory Handler
    ' dirBuf, _ ' Buffer
    ' LenB(dirBuf), _ ' Byte Length of Buffer
    ' True, _ ' Watch sub tree
    ' FILE_NOTIFY_CHANGE_FILE_NAME, _ ' Notification Filter
    ' nUsed, _
    ' VarPtr(oLap), _ ' Pointer to OverLapped
    ' 0&)
    
    Call ReadDirectoryChangesW(hDir, _
                            VarPtr(buf(0)), _
                              UBound(buf) + 1, _
                              True, _
                               FILE_NOTIFY_CHANGE_LAST_WRITE, _
                              nUsed, _
                              VarPtr(oLap), _
                              0&)
    
    'Just for completeness, if you want to use Sync mode
    'Sync mode is called this way
    'Call ReadDirectoryChangesW(hDir, _
    ' VarPtr(buf(0)), _
    ' UBound(buf) + 1, _
    ' False, _
    ' FILE_NOTIFY_CHANGE_FILE_NAME, _
    ' nUsed, _
    ' 0&, _
    ' 0&)
End Sub




Sub Check()


    
    
    fstop = False
    

                                  
                                  
    Dim bstr As String 'string for display of filename
    Dim pos As Long 'location of file entry in buffer buf
    Dim ret As Long 'return value of wait operation
    

     pos = 0
     'wait for hEvent to be unset
     ret = WaitForSingleObject(hEvent, 100)
     If ret <> TIME_OUT Then
        'get the first entry in buf
        CopyMemory VarPtr(dirBuf), VarPtr(buf(pos)), Len(dirBuf)
        'uncomment if you want to see details of dirBuf
        'Debug.Print dirBuf.dwFileNameLength & " " & _
        ' dirBuf.dwNextEntryOffset
        bstr = dirBuf.wcFileName
        
        If dirBuf.dwAction Then RaiseEvent FileAction(bstr, dirBuf.dwAction)
        
        Select Case dirBuf.dwAction
         Case FILE_ACTION_ADDED: Debug.Print "ADDED ";
         'Not mointoring File modification
         Case FILE_ACTION_MODIFIED: Debug.Print "MODIFIED ";
         Case FILE_ACTION_REMOVED: Debug.Print "REMOVED ";
         Case FILE_ACTION_RENAMED_NEW_NAME: Debug.Print "RENAME NEW ";
         Case FILE_ACTION.FILE_ACTION_RENAMED_OLD_NAME:
                          Debug.Print "RENAME OLD ";
                          
        End Select
          
        'locate the filename returned
        'assigned a unicode byte array to a string
        
'        bstr = Left(bstr, dirBuf.dwFileNameLength / 2)
        Debug.Print bstr


        While dirBuf.dwNextEntryOffset <> 0 ' 0 this indicate last entry
          'get next entry in buf
          pos = pos + dirBuf.dwNextEntryOffset
          CopyMemory VarPtr(dirBuf), VarPtr(buf(pos)), Len(dirBuf)
          'uncomment if you want to see details of dirBuf
          'Debug.Print dirBuf.dwFileNameLength & " " & _
           ' dirBuf.dwNextEntryOffset
            bstr = dirBuf.wcFileName
          If dirBuf.dwAction Then RaiseEvent FileAction(bstr, dirBuf.dwAction)
          
          Select Case dirBuf.dwAction
           Case FILE_ACTION_ADDED: Debug.Print "ADDED ";
           'Not mointoring File modification
           Case FILE_ACTION_MODIFIED: Debug.Print "MODIFIED ";
           Case FILE_ACTION_REMOVED: Debug.Print "REMOVED ";
           Case FILE_ACTION_RENAMED_NEW_NAME: Debug.Print "RENAME NEW ";
           Case FILE_ACTION.FILE_ACTION_RENAMED_OLD_NAME:
                         Debug.Print "RENAME OLD ";
          End Select
            
'            bstr = Left(bstr, dirBuf.dwFileNameLength / 2)
            Debug.Print bstr
        Wend
          
          
        'reset the event mointering and repeat the whole process
        ResetEvent hEvent
        Call ReadDirectoryChangesW(hDir, _
                VarPtr(buf(0)), _
                UBound(buf) + 1, _
                True, _
                FILE_NOTIFY_CHANGE_FILE_NAME, _
                nUsed, VarPtr(oLap), 0&)
     End If
     
     DoEvents
       
End Sub




