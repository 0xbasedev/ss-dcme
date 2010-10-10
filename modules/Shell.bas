Attribute VB_Name = "mdlShell"

Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_INVOKEIDLIST = &HC
Option Explicit


Public Const SEE_MASK_FLAG_NO_UI = &H400

Public Const INFINITE = &HFFFF 'Infinite Wait Time
Public Const SW_NORMAL = 1

Public Const WAIT_ABANDONED = &H80
Public Const WAIT_OBJECT_0 = &H0
Public Const WAIT_TIMEOUT = &H102
Public Const WAIT_FAILED = &HFFFFFFFF


Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' Optional Fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200         ' get display name
Private Const SHGFI_EXETYPE = &H2000           ' return exe type
Private Const SHGFI_LARGEICON = &H0           ' get large icon
Private Const SHGFI_SHELLICONSIZE = &H4         ' get shell size icon
Private Const SHGFI_SMALLICON = &H1           ' get small icon
Private Const SHGFI_SYSICONINDEX = &H4000        ' get system icondex
Private Const SHGFI_TYPENAME = &H400           ' get type name
Private Const ILD_BLEND50 = &H4
Private Const ILD_BLEND25 = &H2
Private Const ILD_TRANSPARENT = &H1
Private Const CLR_NONE = &HFFFFFFFF
Private Const CLR_DEFAULT = &HFF000000
Private Type SHFILEINFO
    hIcon As Long           ' : icon
    iIcon As Long     ' : icondex
    dwAttributes As Long        ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80     ' : type name
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long




'Public Sub ExecAndWait(ByVal hwnd As Long, ByVal ProgramPath As String)
'
'    Dim SEI As SHELLEXECUTEINFO
'
'    'Filling SEI structure
'    With SEI
'        .cbSize = Len(SEI) 'Bytes of the structure
'        .fMask = SEE_MASK_NOCLOSEPROCESS 'I need hProcess to be retrieved
'        .lpFile = ProgramPath 'Program Path
'        .lpVerb = "Open" 'Action to do
'        .nShow = SW_NORMAL 'How the program will be showed
'        .lpDirectory = Left(ProgramPath, (InStrRev(ProgramPath, "\")) - 1)
'        .hwnd = hwnd 'Window Handle
'
'    End With
'
'    ShellExecuteEx SEI 'Execute the program hProcess recives the Process Handle used next.
'
'    WaitForSingleObject SEI.hProcess, INFINITE  'Here wait until close the program
'
'    CloseHandle hProcess
'
'End Sub

Global CancelWait As Boolean


'Public Function EditAndWait(ByRef ownerForm As Form, ByVal hWnd As Long, ByVal filepath As String) As Boolean
'    Dim SEI As SHELLEXECUTEINFO
'
'    'Filling SEI structure
'    With SEI
'        .cbSize = Len(SEI) 'Bytes of the structure
'        .fMask = SEE_MASK_NOCLOSEPROCESS 'I need hProcess to be retrieved
'        .lpFile = filepath 'Program Path
'        .lpVerb = "Open" 'Action to do
'        .nShow = SW_NORMAL 'How the program will be showed
'        .lpDirectory = Left(filepath, (InStrRev(filepath, "\")) - 1)
'        .hWnd = hWnd 'Window Handle
'
'    End With
'
'    ShellExecuteEx SEI 'Execute the program hProcess recives the Process Handle used next.
'
'    'Wait for either the program to close, or modifications on the file
'    Dim waitret As Long
'
'    CancelWait = False 'Will be set to true by clicking 'Cancel' on the dialog
'    dlgCancelWait.show , ownerForm
'
'    ownerForm.Enabled = False
'
'    BringWindowToTop dlgCancelWait.hWnd
'
'    Call dlgCancelWait.modifcheck.InitCheck(filepath)
'
'    Do
'        If edit Then
'            'Check if the file is modified
'            Call dlgCancelWait.modifcheck.Check
'
'        End If
'
'        waitret = WaitForSingleObject(SEI.hProcess, 25)
'        DoEvents
'
'
'    Loop Until waitret <> WAIT_TIMEOUT Or CancelWait
'
'    Call dlgCancelWait.modifcheck.StopChecking
'
'    Unload dlgCancelWait
'
'    ownerForm.Enabled = True
'
'
'    CloseHandle SEI.hProcess
'
'    BringWindowToTop ownerForm.hWnd
'End Function


Public Function ExecAndWait(ByRef ownerForm As Form, ByVal hWnd As Long, ByVal ProgramPath As String, ByVal parameters As String) As Boolean
    Dim SEI As SHELLEXECUTEINFO
    
    'Filling SEI structure
    With SEI
        .cbSize = Len(SEI) 'Bytes of the structure
        .fMask = SEE_MASK_NOCLOSEPROCESS 'I need hProcess to be retrieved
        .lpFile = """" & ProgramPath & """" 'Program Path
        .lpVerb = "Open" 'Action to do
        .nShow = SW_NORMAL 'How the program will be showed
        .lpDirectory = Left(ProgramPath, (InStrRev(ProgramPath, "\")) - 1)
        .hWnd = hWnd 'Window Handle
        
        .lpParameters = """" & parameters & """"
    End With
    
    
    
    'Wait for either the program to close, or modifications on the file
    Dim waitret As Long
    
    CancelWait = False 'Will be set to true by clicking 'Cancel' on the dialog
    dlgCancelWait.show , ownerForm
    
    ownerForm.Enabled = False
    
    BringWindowToTop dlgCancelWait.hWnd
    
    ShellExecuteEx SEI 'Execute the program hProcess recives the Process Handle used next.
    
    Do

        waitret = WaitForSingleObject(SEI.hProcess, 25)
        DoEvents
        
        
    Loop Until waitret <> WAIT_TIMEOUT Or CancelWait
    
    Unload dlgCancelWait
    
    ownerForm.Enabled = True
    
    CloseHandle SEI.hProcess
    
    BringWindowToTop ownerForm.hWnd
    
End Function




Function EditImage(ByRef ownerForm As Form, ByVal hDC, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, locksize As Boolean, Optional ByRef targetPic As PictureBox = Nothing) As Boolean
'          Dim ret As Boolean
'
'10        Load frmTileEditor
'
'20        Call frmTileEditor.setParent(srcPic, X, Y, width, height, ByVal VarPtr(ret))
'
'30        Call frmTileEditor.InitImage
'
'40        frmTileEditor.show vbModal, srcPic.parent
'
'50        EditImage = ret
    
    'Copy the portion of image on the temporary picturebox
    
    Dim fname As String
    CreateDir Directory_Temp
    fname = Directory_Temp & "\tmp_bmpedit" & GetTickCount & ".bmp"
    
    With frmGeneral.pictemp
        .Cls
        
        .width = width
        .height = height
        
        BitBlt .hDC, 0, 0, width, height, hDC, X, Y, vbSrcCopy
    
        .Picture = .Image
        
        Call SavePicture(.Picture, fname)

        Dim imgeditor As String
        
        imgeditor = GetSetting("ImageEditor", GetSystemImageEditor)

        ExecAndWait ownerForm, ownerForm.hWnd, imgeditor, fname
        
        If FileExists(fname) Then
    
            frmGeneral.pictemp.Picture = LoadPicture("")
            Call LoadPic(frmGeneral.pictemp, fname)
            .AutoSize = True
            .AutoSize = False
            
            If locksize Then
                If .width <> width Or .height <> height Then
                    MessageBox "Edited image size does not match the original (" & width & "x" & height & ")", vbExclamation + vbOKOnly
                    Call DeleteFile(fname)
                    EditImage = False
                    Exit Function
                End If
            End If
            
            
            If Not targetPic Is Nothing Then
                targetPic.width = .width
                targetPic.height = .height
'                targetPic.Cls
            End If
            
            BitBlt hDC, X, Y, .width, .height, .hDC, 0, 0, vbSrcCopy
            
            If Not targetPic Is Nothing Then targetPic.Refresh
            
            .Picture = LoadPicture("")
            
            EditImage = True
            

            
            Call DeleteFile(fname)
        End If
    
    End With
End Function






Function DrawFileIconOn(filename As String, hdcDst As Long, X As Long, Y As Long) As Boolean
    
    Dim hImage As Long, udtFI As SHFILEINFO
    'set the graphics mode of form1 to persistent
    
    'get the handle of the system image list that contains the large icon images
    hImage = SHGetFileInfo(filename, ByVal 0&, udtFI, Len(udtFI), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
    
    If hImage <> 0 Then
        'draw the icon (normal)
        ImageList_Draw hImage, udtFI.iIcon, hdcDst, X, Y, ILD_TRANSPARENT
        DrawFileIconOn = True
    Else
        DrawFileIconOn = False
    End If
    
'    DestroyIcon hImage
    
End Function
