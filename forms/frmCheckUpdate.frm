VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Checking for updates"
   ClientHeight    =   6330
   ClientLeft      =   300
   ClientTop       =   585
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   1  'CenterOwner
   Begin DCME.cProgressBar progress 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   4250
      DisplayMode     =   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   6015
      Begin VB.ComboBox cmbUpdateDelay 
         Height          =   315
         ItemData        =   "frmCheckUpdate.frx":0000
         Left            =   3000
         List            =   "frmCheckUpdate.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   180
         Width           =   2055
      End
      Begin VB.CheckBox chkAutoUpdate 
         Caption         =   "Automatically Check For Updates"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox txtDetails 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2880
      Width           =   6015
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   4680
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton cmdUpdateNow 
      Caption         =   "Update Now"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdRecheck 
      Caption         =   "Check For Updates"
      Default         =   -1  'True
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Download Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblGlobalProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   5895
   End
   Begin VB.Label lblCurrentVersion 
      Caption         =   "Current version: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblstatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblRetrievingUpdateInfo 
      Caption         =   "Retrieving version information..."
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblconnect 
      Caption         =   "Connecting to server..."
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image imgarrow 
      Height          =   330
      Left            =   120
      Picture         =   "frmCheckUpdate.frx":004F
      Top             =   660
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim updateVersion As Long
Dim currentVersion As Long

Dim updateVersionStr As String
Dim currentVersionStr As String

Dim FilesToUpdate() As tFileToUpdate
Dim NumberOfFiles As Integer

Dim Changes() As tChange
Dim NumberOfChanges As Integer


Const DETAILS_SEPARATOR = "-------------------------------"

Private Sub chkAutoUpdate_Click()
    If chkAutoUpdate.value = vbChecked Then
        cmbUpdateDelay.Enabled = True
    Else
        cmbUpdateDelay.Enabled = False
    End If
End Sub

Private Sub cmdRecheck_Click()
    'Rechecks if an update is available
    CheckUpdate
End Sub

Private Sub cmdUpdate_Click()
    'Download and install the update
    
    Dim exePath As String

    cmdRecheck.Enabled = False
    cmdUpdate.Enabled = False
    lblstatus.Caption = "Downloading version " & updateVersionStr & " ..."
    
    imgarrow.Top = lblRetrievingUpdateInfo.Top - 4
    
    Dim i As Integer
    
    'calculate total size
    Dim totalsize As Long
    Dim doneSize As Long
    Dim tickStart As Long
    
    'calculate total size
    totalsize = 0
    doneSize = 0
    For i = 0 To NumberOfFiles - 1
        totalsize = totalsize + FilesToUpdate(i).filesize
    Next
    
    progress.value = 0
    progress.Max = totalsize
    progress.displaymode = dBytes
    
    tickStart = GetTickCount
    For i = 0 To NumberOfFiles - 1
        AppendDetails "Downloading file " & i + 1 & "/" & NumberOfFiles & ": " & GetFileTitle(FilesToUpdate(i).localpath)
        lblRetrievingUpdateInfo.Caption = "Downloading " & GetFileTitle(FilesToUpdate(i).localpath)
        
        AddDebug "+++ Downloading from " & FilesToUpdate(i).url & " to " & FilesToUpdate(i).localpath
        
        'download each file
        Call HTTPDownloadFile(Inet, FilesToUpdate(i), totalsize, doneSize, tickStart)

        DoEvents
        
        If FilesToUpdate(i).sfx Then
            'extract files from self-extracting archive
            
            If FileExists(FilesToUpdate(i).localpath) Then
                AddDebug "+++ Extracting files from " & FilesToUpdate(i).localpath
                AppendDetails "Unpacking archive content..."
'                Shell FilesToUpdate(i).localpath, vbHide
                ShellExecute 0&, vbNullString, FilesToUpdate(i).localpath, vbNullString, GetPathTo(FilesToUpdate(i).localpath), vbHide
            Else
                AddDebug "+++ Could not extract files from " & FilesToUpdate(i).localpath
                AppendDetails "Error while unpacking archive content"
            End If
        End If
        
        If LCase(GetFileTitle(FilesToUpdate(i).localpath)) = "dcmeupdate.exe" Or LCase(GetFileTitle(FilesToUpdate(i).localpath)) = "dcme.exe" Then
            exePath = FilesToUpdate(i).localpath
        End If
        
        doneSize = doneSize + FilesToUpdate(i).filesize
    Next

    progress.value = totalsize
    
    AppendDetails "Update successful. The update will be completed when you close DCME."
    AppendDetails "Click ''Update Now'' or restart DCME for changes to take effect now."
    AppendDetails DETAILS_SEPARATOR
    AddDebug "+++ Update files successfully downloaded."

    lblRetrievingUpdateInfo.Caption = "Download successful."
    imgarrow.Top = lblstatus.Top - 4
    lblstatus.Caption = "DCME version " & updateVersionStr & " ready"

    DoEvents


    cmdUpdateNow.Enabled = True
    cmdRecheck.Enabled = True
    
    If exePath <> "" Then
        frmGeneral.updateready = True
        frmGeneral.updatefilepath = exePath
        
        If MessageBox("DCME " & updateVersionStr & " was successfully downloaded. Do you wish to install the update now? You will be prompted to save your maps.", vbYesNo + vbQuestion, "Update downloaded") = vbYes Then
            Unload Me
            Unload frmGeneral
        End If
    Else
        MessageBox "DCME " & updateVersionStr & " was successfully downloaded. Please restart DCME for the changes to take effect."
        'Only external files were updated
    End If

End Sub


Private Sub cmdUpdateNow_Click()
    'Closes DCME for update (self extracting archive will start DCME)
    Unload Me
    Unload frmGeneral
End Sub

Private Sub Form_Load()
    'Update gui with current version and check if an update is available
    Set Me.Icon = frmGeneral.Icon

    frmGeneral.updateformloaded = True
    
    currentVersion = CLng(App.Major) * 100000 + CLng(App.Minor) * 1000 + CLng(App.Revision) * 1
    currentVersionStr = App.Major & "." & App.Minor & "." & App.Revision
    
    AppendDetails "Current version: " & currentVersionStr
    
    chkAutoUpdate.value = GetSetting("AutoUpdate", 1)
    cmbUpdateDelay.ListIndex = GetSetting("AutoUpdateDelay", 2)

    If chkAutoUpdate.value = vbChecked Then
        cmbUpdateDelay.Enabled = True
    Else
        cmbUpdateDelay.Enabled = False
    End If


    lblconnect.visible = False
    imgarrow.visible = False

    lblCurrentVersion.Caption = "Current version: " & currentVersionStr
    Me.visible = Not frmGeneral.quickupdate
    CheckUpdate

End Sub

Private Sub AppendDetails(Text As String)
    'Add a line to the details textbox
    txtDetails.Text = txtDetails.Text & Text & vbNewLine
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Check if the update is still running, and ask for abort
    If Inet.StillExecuting Then
        If MessageBox("Cancel update?", vbYesNo + vbQuestion) = vbYes Then
            Inet.Cancel
            Unload Me
        Else
            Cancel = True
        End If
    End If
End Sub

Sub CheckUpdate()
    'Checks for update of any files more recent than the current version in the dcmeupdate.txt
    progress.value = 0
    progress.Max = 100
    
    progress.DisplayDecimals = 0
    progress.displaymode = dPercentage
    
    
    Call SetSetting("AutoUpdateLast", CStr(Day(Now) & "-" & Month(Now) & "-" & Year(Now)))

    cmdUpdate.Enabled = False
    cmdRecheck.Enabled = False
    lblstatus.Caption = "Searching for updates..."
    lblRetrievingUpdateInfo.visible = False

    lblconnect.visible = True
    imgarrow.visible = True

    'reset labels
    lblconnect.Caption = "Connecting to server..."
    lblRetrievingUpdateInfo.Caption = "Retrieving version information..."
    lblRetrievingUpdateInfo.visible = True

    imgarrow.Top = lblconnect.Top - 4

    

    Dim str As String
    Dim updateinfofile As String
    updateinfofile = GetSetting("UpdateURL", DEFAULT_UPDATE_URL)
    
    If GetExtension(updateinfofile) <> "txt" And GetExtension(updateinfofile) <> "ini" Then
        updateinfofile = updateinfofile & "dcmeupdate.txt"
    End If
    
    AppendDetails "Searching for updates..."
    AddDebug "+++ Searching for updates at " & updateinfofile & "..."
    
    'retrieve the update list file
    str = GetHTTPFile(Inet, updateinfofile)
        
    progress.value = 20
    
    lblconnect.Caption = "Connected to server."

    
    If InStr(str, "DCME") = 0 Or InStr(str, "was not found on this server") <> 0 Then
        progress.value = 100
        
        'Invalid update info file
        lblRetrievingUpdateInfo.Caption = "Unable to get update information"
        lblstatus.Caption = "Unable to get update info"
        
        If updateinfofile <> DEFAULT_UPDATE_URL Then
            AppendDetails "DCME could not retrieve update information from the server. Please reset the update URL to default in your preferences to solve the problem."
        Else
            AppendDetails "DCME could not retrieve update information from the server. You can get the update manually from http://forums.sscentral.com/index.php?showtopic=10662."
        End If
        
        cmdRecheck.Enabled = True

        AddDebug "+++ Unable to retrieve information from server"

        Exit Sub
    End If
    
    AppendDetails "Connected to server."
    
    imgarrow.Top = lblRetrievingUpdateInfo.Top - 4

    Dim lines() As String
    lines = Split(str, Chr(10))

    Dim i As Integer
    
    NumberOfFiles = 0
    NumberOfChanges = 0
    ReDim FilesToUpdate(0)
    ReDim Changes(0)
    
    updateVersion = currentVersion
    
    For i = LBound(lines) To UBound(lines)
        If lines(i) <> "" Then
            Dim parts() As String
            parts() = Split(lines(i), "::")

            If parts(0) = "file" Or parts(0) = "sfx" Then
                If CLng(parts(1)) > currentVersion Then
                    Call AddFileToUpdate(parts())
                End If
            ElseIf parts(0) = "change" Then
                If CLng(parts(1)) > currentVersion Then
                    Call AddChange(parts())
                End If

            ElseIf parts(0) = "date" Then
            
            End If
        End If
        progress.value = 20 + Int(80 * (i / UBound(lines)))
    Next

    progress.value = 100
        
    updateVersionStr = VersionToString(updateVersion)

    lblRetrievingUpdateInfo.Caption = "Last version available: " & updateVersionStr

    AddDebug "+++ Latest version available: " & updateVersionStr & " (Current version: " & currentVersionStr & ")"
    AppendDetails "Latest version available: " & updateVersionStr
    
    imgarrow.Top = lblRetrievingUpdateInfo.Top - 4
    
    ' format to compare: Mmmrrr (e.g 1.2.53 = 102053)
    If updateVersion > currentVersion Then
        'new version
        lblstatus.Caption = "New Version Available: " & updateVersionStr
        
        'List files to update
        AppendDetails "---"
        AppendDetails "The following files need an update:"
        For i = 0 To NumberOfFiles - 1
            AppendDetails " - " & replace(FilesToUpdate(i).localpath, App.path & "\", "") & " (" & FilesToUpdate(i).description & " version " & FilesToUpdate(i).version \ 100000 & "." & (FilesToUpdate(i).version Mod 100000) \ 1000 & "." & FilesToUpdate(i).version Mod 1000 & ") - " & FilesToUpdate(i).filesize \ 1024 & "KB"
        Next
        
        'List changes
        AppendDetails "---"
        AppendDetails "Latest changes available :"
        AppendDetails updateVersionStr & " :"
        
        Dim lastChangeVersion As Long
        lastChangeVersion = updateVersion
        
        For i = 0 To NumberOfChanges - 1
            If Changes(i).version <= updateVersion Then 'Don't show updates that are in the file, but not uploaded yet
                If Changes(i).version < lastChangeVersion Then
                    AppendDetails VersionToString(Changes(i).version) & " :"
                    lastChangeVersion = Changes(i).version
                End If
                AppendDetails " - " & Changes(i).description
            End If
        Next
        
        lblstatus.Caption = "New version available: " & updateVersionStr
        
        If updateVersion - currentVersion >= 100000 Then

            If frmGeneral.quickupdate Then
                If MessageBox("A major update of DCME is available, do you wish to download it?", vbYesNo + vbExclamation, "Major update available (version " & updateVersionStr & ")") = vbYes Then
                    cmdUpdate_Click
                End If
            End If
        ElseIf updateVersion - currentVersion >= 1000 Then
            
            If frmGeneral.quickupdate Then
                If MessageBox("An important update of DCME is available, do you wish to download it?", vbYesNo + vbQuestion, "Important update available (version " & updateVersionStr & ")") = vbYes Then
                    cmdUpdate_Click
                End If
            End If

        ElseIf updateVersion - currentVersion >= 1 Then

            If frmGeneral.quickupdate Then
                If MessageBox("A new version of DCME is available, do you wish to download the update?", vbYesNo + vbQuestion, "Update available (version " & updateVersionStr & ")") = vbYes Then
                    cmdUpdate_Click
                End If
            End If

        End If

        cmdUpdate.Enabled = True

    Else
        lblstatus.Caption = "No Updates Available"
        AppendDetails "No new updates available"
    End If
    
    
    AppendDetails DETAILS_SEPARATOR
    
    
    DoEvents
    
    imgarrow.Top = lblstatus.Top - 4
    
    cmdRecheck.Enabled = True

End Sub

Private Sub AddFileToUpdate(info() As String)
    If UBound(info) < 4 Then Exit Sub
    

    
    Dim fileversion As Long
    Dim filepath As String
    
    fileversion = CLng(info(1))
    filepath = App.path & "\" & info(2)
    
    'Check if same file already exists
    Dim i As Integer
    For i = 0 To NumberOfFiles
        If FilesToUpdate(i).localpath = filepath Then
            'File is already there...
            If FilesToUpdate(i).version > fileversion Then
                'File is already there, and is a more recent version
                Exit Sub
            End If
        End If
    Next
    
    'Redim for new file if necessary
    If NumberOfFiles > UBound(FilesToUpdate) Then
        ReDim Preserve FilesToUpdate(NumberOfFiles + 5)
    End If
    
    FilesToUpdate(NumberOfFiles).version = fileversion
    FilesToUpdate(NumberOfFiles).localpath = filepath
    FilesToUpdate(NumberOfFiles).url = info(3)
    FilesToUpdate(NumberOfFiles).description = info(4)
    
    FilesToUpdate(NumberOfFiles).sfx = (info(0) = "sfx")
    
    FilesToUpdate(NumberOfFiles).filesize = ReadFileSize(Inet, FilesToUpdate(NumberOfFiles).url)
    
    If updateVersion < CLng(info(1)) Then
        updateVersion = CLng(info(1))
    End If
    
    NumberOfFiles = NumberOfFiles + 1
    
End Sub

Private Sub AddChange(info() As String)
    If UBound(info) < 2 Then Exit Sub
    
    If NumberOfChanges > UBound(Changes) Then
        ReDim Preserve Changes(NumberOfChanges + 5)
    End If
    
    
    Changes(NumberOfChanges).version = CLng(info(1))
    Changes(NumberOfChanges).description = info(2)
    
    NumberOfChanges = NumberOfChanges + 1
    
End Sub

'GetHTTPFile **************************************
'
' Return contents of an file from a given HTTP URL
'
' Requires: Reference to Internet Transfer control
'           Valid URL that returns textual data
'
' Author: Richard Bowman
' Date:   02/17/2005
'
Function GetHTTPFile(ctlInet As Inet, ByVal sURL As String) As String
    On Error GoTo GetHTTPFile_Error

' cancel old operations
    ctlInet.Cancel
    ' set protocol to HTTP
    ctlInet.Protocol = icHTTP
    ' get the page
    GetHTTPFile = ctlInet.OpenURL(sURL)
    
    On Error GoTo 0
    Exit Function
GetHTTPFile_Error:
    If Err.Number <> ERR_MSINET_REQUESTTIMEDOUT Then
        HandleError Err, "frmCheckUpdate.GetHTTPFile", False
    End If
    GetHTTPFile = ""
End Function




Private Sub Form_Unload(Cancel As Integer)
'save update settings
    Call SetSetting("AutoUpdate", chkAutoUpdate.value)
    Call SetSetting("AutoUpdateDelay", cmbUpdateDelay.ListIndex)

    Call settings.SaveSettings

    frmGeneral.updateformloaded = False
End Sub

Private Sub inet_StateChanged(ByVal State As Integer)
    If State = icConnecting Then
        lblconnect.Caption = "Connecting to server..."
        imgarrow.Top = lblconnect.Top - 4
        DoEvents
    ElseIf State = icReceivingResponse Then
        If imgarrow.Top <> lblRetrievingUpdateInfo.Top - 4 Then
            imgarrow.Top = lblRetrievingUpdateInfo.Top - 4
        End If
        lblconnect.Caption = "Connected to server."
        lblRetrievingUpdateInfo.visible = True
        DoEvents
    Else
        DoEvents
    End If

End Sub


Private Function ReadFileSize(ByRef http As Inet, ByVal url As String) As Long

    Dim strHeader As String
    
    With http
        .Protocol = icHTTP
        .url = url
        .Execute , "GET", , "Range: bytes=0-" & vbCrLf
    
        While .StillExecuting
            DoEvents
        Wend
        
    End With
    
    strHeader = http.GetHeader("Content-Length")
    
    http.Cancel
    ReadFileSize = val(strHeader)
End Function

Private Sub HTTPDownloadFile(ByRef http As Inet, fileInfo As tFileToUpdate, totalsize As Long, doneSize As Long, tickStart As Long)
    
    On Error GoTo HTTPDownloadFile_Error
    
    Const Chunk_Size As Long = 1024
    Dim f As Integer
    Dim strHeader As String
    Dim b() As Byte
    Dim tmpstr As String
    Dim fullstr As String
    fullstr = ""
    Dim lngBytesReceived As Long
    Dim lngFileLength As Long
    
    DoEvents
    With http
        .Protocol = icHTTP
        .url = fileInfo.url
        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
        While .StillExecuting
            DoEvents
        Wend
        
    End With
    
    lngFileLength = fileInfo.filesize
    
    DoEvents
    
    lngBytesReceived = 0
    f = FreeFile()
    
    'create needed folder
    CreateDir GetPathTo(fileInfo.localpath)
    Dim lasttick As Long
    Dim strprogress As String
    
    If FileExists(fileInfo.localpath) Then Kill (fileInfo.localpath)
    
    If IsTextFile(fileInfo.localpath) Then
        Do
            tmpstr = http.GetChunk(Chunk_Size, icString)
            fullstr = fullstr & tmpstr
            lngBytesReceived = lngBytesReceived + Len(tmpstr)
            
            strprogress = "Downloading " & GetFileTitle(fileInfo.localpath) & " " & IIf(fileInfo.filesize > 0, " (" & Int(lngBytesReceived / (fileInfo.filesize + 1) * 100) & "% of " & IIf(fileInfo.filesize > 1048576, Format(fileInfo.filesize / 1024 / 1024, "0.00") & "MB)", fileInfo.filesize \ 1024 & "KB)"), "")
            If lblRetrievingUpdateInfo.Caption <> strprogress Then
                lblRetrievingUpdateInfo.Caption = strprogress
            End If
            
            If totalsize > 0 Then
                progress.value = doneSize + lngBytesReceived
                
                If GetTickCount - lasttick > 200 Or lngBytesReceived >= lngFileLength Then 'Update rate
                    lblGlobalProgress.Caption = "Total Progress: " & Int(((doneSize + lngBytesReceived) / totalsize) * 100) & "% " & IIf(GetTickCount - tickStart > 0, " @ " & Format((doneSize + lngBytesReceived) / (GetTickCount - tickStart), "#") & "KB/s", "")
                    lasttick = GetTickCount
                End If
                
            End If
        Loop While Len(tmpstr) > 0
        
        fullstr = replace(fullstr, Chr$(10), vbNewLine)
        Open fileInfo.localpath For Output As #f
            Print #f, fullstr
        Close #f
    Else
        Open fileInfo.localpath For Binary Access Write As #f
        Do
            b = http.GetChunk(Chunk_Size, icByteArray)
            Put #f, , b
            lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1
            
            strprogress = "Downloading " & GetFileTitle(fileInfo.localpath) & " " & IIf(fileInfo.filesize > 0, " (" & Int(lngBytesReceived / (fileInfo.filesize + 1) * 100) & "% of " & IIf(fileInfo.filesize > 1048576, Format(fileInfo.filesize / 1024 / 1024, "0.00") & "MB)", fileInfo.filesize \ 1024 & "KB)"), "")
            If lblRetrievingUpdateInfo.Caption <> strprogress Then
                lblRetrievingUpdateInfo.Caption = strprogress
            End If
            
            If totalsize > 0 Then
                progress.value = doneSize + lngBytesReceived
                
                If GetTickCount - lasttick > 200 Or lngBytesReceived >= lngFileLength Then 'Update rate
                    lblGlobalProgress.Caption = "Total Progress: " & Int(((doneSize + lngBytesReceived) / totalsize) * 100) & "% " & IIf(GetTickCount - tickStart > 0, " @ " & Format((doneSize + lngBytesReceived) / (GetTickCount - tickStart), "#") & "KB/s", "")
                    lasttick = GetTickCount
                End If
                
            End If
        Loop While UBound(b, 1) > 0
        Close #f
    End If
    
    
    lngFileLength = 0
    lngBytesReceived = 0

    On Error GoTo 0
    Exit Sub
    
HTTPDownloadFile_Error:
    HandleError Err, "frmCheckUpdate.HTTPDownloadFile", False
    
End Sub



Private Sub txtDetails_Change()
    txtDetails.selstart = Len(txtDetails.Text)
End Sub
