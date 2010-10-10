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
10        If chkAutoUpdate.value = vbChecked Then
20            cmbUpdateDelay.Enabled = True
30        Else
40            cmbUpdateDelay.Enabled = False
50        End If
End Sub

Private Sub cmdRecheck_Click()
          'Rechecks if an update is available
10        CheckUpdate
End Sub

Private Sub cmdUpdate_Click()
          'Download and install the update
          
          Dim exePath As String

10        cmdRecheck.Enabled = False
20        cmdUpdate.Enabled = False
30        lblstatus.Caption = "Downloading version " & updateVersionStr & " ..."
          
40        imgarrow.Top = lblRetrievingUpdateInfo.Top - 4
          
          Dim i As Integer
          
          'calculate total size
          Dim totalsize As Long
          Dim doneSize As Long
          Dim tickStart As Long
          
          'calculate total size
50        totalsize = 0
60        doneSize = 0
70        For i = 0 To NumberOfFiles - 1
80            totalsize = totalsize + FilesToUpdate(i).filesize
90        Next
          
100       progress.value = 0
110       progress.Max = totalsize
120       progress.displaymode = dBytes
          
130       tickStart = GetTickCount
140       For i = 0 To NumberOfFiles - 1
150           AppendDetails "Downloading file " & i + 1 & "/" & NumberOfFiles & ": " & GetFileTitle(FilesToUpdate(i).localpath)
160           lblRetrievingUpdateInfo.Caption = "Downloading " & GetFileTitle(FilesToUpdate(i).localpath)
              
170           AddDebug "+++ Downloading from " & FilesToUpdate(i).url & " to " & FilesToUpdate(i).localpath
              
              'download each file
180           Call HTTPDownloadFile(Inet, FilesToUpdate(i), totalsize, doneSize, tickStart)

190           DoEvents
              
200           If FilesToUpdate(i).sfx Then
                  'extract files from self-extracting archive
                  
210               If FileExists(FilesToUpdate(i).localpath) Then
220                   AddDebug "+++ Extracting files from " & FilesToUpdate(i).localpath
230                   AppendDetails "Unpacking archive content..."
      '                Shell FilesToUpdate(i).localpath, vbHide
240                   ShellExecute 0&, vbNullString, FilesToUpdate(i).localpath, vbNullString, GetPathTo(FilesToUpdate(i).localpath), vbHide
250               Else
260                   AddDebug "+++ Could not extract files from " & FilesToUpdate(i).localpath
270                   AppendDetails "Error while unpacking archive content"
280               End If
290           End If
              
300           If LCase(GetFileTitle(FilesToUpdate(i).localpath)) = "dcmeupdate.exe" Or LCase(GetFileTitle(FilesToUpdate(i).localpath)) = "dcme.exe" Then
310               exePath = FilesToUpdate(i).localpath
320           End If
              
330           doneSize = doneSize + FilesToUpdate(i).filesize
340       Next

350       progress.value = totalsize
          
360       AppendDetails "Update successful. The update will be completed when you close DCME."
370       AppendDetails "Click ''Update Now'' or restart DCME for changes to take effect now."
380       AppendDetails DETAILS_SEPARATOR
390       AddDebug "+++ Update files successfully downloaded."

400       lblRetrievingUpdateInfo.Caption = "Download successful."
410       imgarrow.Top = lblstatus.Top - 4
420       lblstatus.Caption = "DCME version " & updateVersionStr & " ready"

430       DoEvents


440       cmdUpdateNow.Enabled = True
450       cmdRecheck.Enabled = True
          
460       If exePath <> "" Then
470           frmGeneral.updateready = True
480           frmGeneral.updatefilepath = exePath
              
490           If MessageBox("DCME " & updateVersionStr & " was successfully downloaded. Do you wish to install the update now? You will be prompted to save your maps.", vbYesNo + vbQuestion, "Update downloaded") = vbYes Then
500               Unload Me
510               Unload frmGeneral
520           End If
530       Else
540           MessageBox "DCME " & updateVersionStr & " was successfully downloaded. Please restart DCME for the changes to take effect."
              'Only external files were updated
550       End If

End Sub


Private Sub cmdUpdateNow_Click()
          'Closes DCME for update (self extracting archive will start DCME)
10        Unload Me
20        Unload frmGeneral
End Sub

Private Sub Form_Load()
          'Update gui with current version and check if an update is available
10        Set Me.Icon = frmGeneral.Icon

20        frmGeneral.updateformloaded = True
          
30        currentVersion = CLng(App.Major) * 100000 + CLng(App.Minor) * 1000 + CLng(App.Revision) * 1
40        currentVersionStr = App.Major & "." & App.Minor & "." & App.Revision
          
50        AppendDetails "Current version: " & currentVersionStr
          
60        chkAutoUpdate.value = GetSetting("AutoUpdate", 1)
70        cmbUpdateDelay.ListIndex = GetSetting("AutoUpdateDelay", 2)

80        If chkAutoUpdate.value = vbChecked Then
90            cmbUpdateDelay.Enabled = True
100       Else
110           cmbUpdateDelay.Enabled = False
120       End If


130       lblconnect.visible = False
140       imgarrow.visible = False

150       lblCurrentVersion.Caption = "Current version: " & currentVersionStr
160       Me.visible = Not frmGeneral.quickupdate
170       CheckUpdate

End Sub

Private Sub AppendDetails(Text As String)
          'Add a line to the details textbox
10        txtDetails.Text = txtDetails.Text & Text & vbNewLine
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
          'Check if the update is still running, and ask for abort
10        If Inet.StillExecuting Then
20            If MessageBox("Cancel update?", vbYesNo + vbQuestion) = vbYes Then
30                Inet.Cancel
40                Unload Me
50            Else
60                Cancel = True
70            End If
80        End If
End Sub

Sub CheckUpdate()
          'Checks for update of any files more recent than the current version in the dcmeupdate.txt
10        progress.value = 0
20        progress.Max = 100
          
30        progress.DisplayDecimals = 0
40        progress.displaymode = dPercentage
          
          
50        Call SetSetting("AutoUpdateLast", CStr(Day(Now) & "-" & Month(Now) & "-" & Year(Now)))

60        cmdUpdate.Enabled = False
70        cmdRecheck.Enabled = False
80        lblstatus.Caption = "Searching for updates..."
90        lblRetrievingUpdateInfo.visible = False

100       lblconnect.visible = True
110       imgarrow.visible = True

          'reset labels
120       lblconnect.Caption = "Connecting to server..."
130       lblRetrievingUpdateInfo.Caption = "Retrieving version information..."
140       lblRetrievingUpdateInfo.visible = True

150       imgarrow.Top = lblconnect.Top - 4

          

          Dim str As String
          Dim updateinfofile As String
160       updateinfofile = GetSetting("UpdateURL", DEFAULT_UPDATE_URL)
          
170       If GetExtension(updateinfofile) <> "txt" And GetExtension(updateinfofile) <> "ini" Then
180           updateinfofile = updateinfofile & "dcmeupdate.txt"
190       End If
          
200       AppendDetails "Searching for updates..."
210       AddDebug "+++ Searching for updates at " & updateinfofile & "..."
          
          'retrieve the update list file
220       str = GetHTTPFile(Inet, updateinfofile)
              
230       progress.value = 20
          
240       lblconnect.Caption = "Connected to server."

          
250       If InStr(str, "DCME") = 0 Or InStr(str, "was not found on this server") <> 0 Then
260           progress.value = 100
              
              'Invalid update info file
270           lblRetrievingUpdateInfo.Caption = "Unable to get update information"
280           lblstatus.Caption = "Unable to get update info"
              
290           If updateinfofile <> DEFAULT_UPDATE_URL Then
300               AppendDetails "DCME could not retrieve update information from the server. Please reset the update URL to default in your preferences to solve the problem."
310           Else
320               AppendDetails "DCME could not retrieve update information from the server. You can get the update manually from http://forums.sscentral.com/index.php?showtopic=10662."
330           End If
              
340           cmdRecheck.Enabled = True

350           AddDebug "+++ Unable to retrieve information from server"

360           Exit Sub
370       End If
          
380       AppendDetails "Connected to server."
          
390       imgarrow.Top = lblRetrievingUpdateInfo.Top - 4

          Dim lines() As String
400       lines = Split(str, Chr(10))

          Dim i As Integer
          
410       NumberOfFiles = 0
420       NumberOfChanges = 0
430       ReDim FilesToUpdate(0)
440       ReDim Changes(0)
          
450       updateVersion = currentVersion
          
460       For i = LBound(lines) To UBound(lines)
470           If lines(i) <> "" Then
                  Dim parts() As String
480               parts() = Split(lines(i), "::")

490               If parts(0) = "file" Or parts(0) = "sfx" Then
500                   If CLng(parts(1)) > currentVersion Then
510                       Call AddFileToUpdate(parts())
520                   End If
530               ElseIf parts(0) = "change" Then
540                   If CLng(parts(1)) > currentVersion Then
550                       Call AddChange(parts())
560                   End If

570               ElseIf parts(0) = "date" Then
                  
580               End If
590           End If
600           progress.value = 20 + Int(80 * (i / UBound(lines)))
610       Next

620       progress.value = 100
              
630       updateVersionStr = VersionToString(updateVersion)

640       lblRetrievingUpdateInfo.Caption = "Last version available: " & updateVersionStr

650       AddDebug "+++ Latest version available: " & updateVersionStr & " (Current version: " & currentVersionStr & ")"
660       AppendDetails "Latest version available: " & updateVersionStr
          
670       imgarrow.Top = lblRetrievingUpdateInfo.Top - 4
          
          ' format to compare: Mmmrrr (e.g 1.2.53 = 102053)
680       If updateVersion > currentVersion Then
              'new version
690           lblstatus.Caption = "New Version Available: " & updateVersionStr
              
              'List files to update
700           AppendDetails "---"
710           AppendDetails "The following files need an update:"
720           For i = 0 To NumberOfFiles - 1
730               AppendDetails " - " & replace(FilesToUpdate(i).localpath, App.path & "\", "") & " (" & FilesToUpdate(i).description & " version " & FilesToUpdate(i).version \ 100000 & "." & (FilesToUpdate(i).version Mod 100000) \ 1000 & "." & FilesToUpdate(i).version Mod 1000 & ") - " & FilesToUpdate(i).filesize \ 1024 & "KB"
740           Next
              
              'List changes
750           AppendDetails "---"
760           AppendDetails "Latest changes available :"
770           AppendDetails updateVersionStr & " :"
              
              Dim lastChangeVersion As Long
780           lastChangeVersion = updateVersion
              
790           For i = 0 To NumberOfChanges - 1
800               If Changes(i).version <= updateVersion Then 'Don't show updates that are in the file, but not uploaded yet
810                   If Changes(i).version < lastChangeVersion Then
820                       AppendDetails VersionToString(Changes(i).version) & " :"
830                       lastChangeVersion = Changes(i).version
840                   End If
850                   AppendDetails " - " & Changes(i).description
860               End If
870           Next
              
880           lblstatus.Caption = "New version available: " & updateVersionStr
              
890           If updateVersion - currentVersion >= 100000 Then

900               If frmGeneral.quickupdate Then
910                   If MessageBox("A major update of DCME is available, do you wish to download it?", vbYesNo + vbExclamation, "Major update available (version " & updateVersionStr & ")") = vbYes Then
920                       cmdUpdate_Click
930                   End If
940               End If
950           ElseIf updateVersion - currentVersion >= 1000 Then
                  
960               If frmGeneral.quickupdate Then
970                   If MessageBox("An important update of DCME is available, do you wish to download it?", vbYesNo + vbQuestion, "Important update available (version " & updateVersionStr & ")") = vbYes Then
980                       cmdUpdate_Click
990                   End If
1000              End If

1010          ElseIf updateVersion - currentVersion >= 1 Then

1020              If frmGeneral.quickupdate Then
1030                  If MessageBox("A new version of DCME is available, do you wish to download the update?", vbYesNo + vbQuestion, "Update available (version " & updateVersionStr & ")") = vbYes Then
1040                      cmdUpdate_Click
1050                  End If
1060              End If

1070          End If

1080          cmdUpdate.Enabled = True

1090      Else
1100          lblstatus.Caption = "No Updates Available"
1110          AppendDetails "No new updates available"
1120      End If
          
          
1130      AppendDetails DETAILS_SEPARATOR
          
          
1140      DoEvents
          
1150      imgarrow.Top = lblstatus.Top - 4
          
1160      cmdRecheck.Enabled = True

End Sub

Private Sub AddFileToUpdate(info() As String)
10        If UBound(info) < 4 Then Exit Sub
          

          
          Dim fileversion As Long
          Dim filepath As String
          
20        fileversion = CLng(info(1))
30        filepath = App.path & "\" & info(2)
          
          'Check if same file already exists
          Dim i As Integer
40        For i = 0 To NumberOfFiles
50            If FilesToUpdate(i).localpath = filepath Then
                  'File is already there...
60                If FilesToUpdate(i).version > fileversion Then
                      'File is already there, and is a more recent version
70                    Exit Sub
80                End If
90            End If
100       Next
          
          'Redim for new file if necessary
110       If NumberOfFiles > UBound(FilesToUpdate) Then
120           ReDim Preserve FilesToUpdate(NumberOfFiles + 5)
130       End If
          
140       FilesToUpdate(NumberOfFiles).version = fileversion
150       FilesToUpdate(NumberOfFiles).localpath = filepath
160       FilesToUpdate(NumberOfFiles).url = info(3)
170       FilesToUpdate(NumberOfFiles).description = info(4)
          
180       FilesToUpdate(NumberOfFiles).sfx = (info(0) = "sfx")
          
190       FilesToUpdate(NumberOfFiles).filesize = ReadFileSize(Inet, FilesToUpdate(NumberOfFiles).url)
          
200       If updateVersion < CLng(info(1)) Then
210           updateVersion = CLng(info(1))
220       End If
          
230       NumberOfFiles = NumberOfFiles + 1
          
End Sub

Private Sub AddChange(info() As String)
10        If UBound(info) < 2 Then Exit Sub
          
20        If NumberOfChanges > UBound(Changes) Then
30            ReDim Preserve Changes(NumberOfChanges + 5)
40        End If
          
          
50        Changes(NumberOfChanges).version = CLng(info(1))
60        Changes(NumberOfChanges).description = info(2)
          
70        NumberOfChanges = NumberOfChanges + 1
          
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
10        On Error GoTo GetHTTPFile_Error

      ' cancel old operations
20        ctlInet.Cancel
          ' set protocol to HTTP
30        ctlInet.Protocol = icHTTP
          ' get the page
40        GetHTTPFile = ctlInet.OpenURL(sURL)
          
50        On Error GoTo 0
60        Exit Function
GetHTTPFile_Error:
70        If Err.Number <> ERR_MSINET_REQUESTTIMEDOUT Then
80            HandleError Err, "frmCheckUpdate.GetHTTPFile", False
90        End If
100       GetHTTPFile = ""
End Function




Private Sub Form_Unload(Cancel As Integer)
      'save update settings
10        Call SetSetting("AutoUpdate", chkAutoUpdate.value)
20        Call SetSetting("AutoUpdateDelay", cmbUpdateDelay.ListIndex)

30        Call settings.SaveSettings

40        frmGeneral.updateformloaded = False
End Sub

Private Sub inet_StateChanged(ByVal State As Integer)
10        If State = icConnecting Then
20            lblconnect.Caption = "Connecting to server..."
30            imgarrow.Top = lblconnect.Top - 4
40            DoEvents
50        ElseIf State = icReceivingResponse Then
60            If imgarrow.Top <> lblRetrievingUpdateInfo.Top - 4 Then
70                imgarrow.Top = lblRetrievingUpdateInfo.Top - 4
80            End If
90            lblconnect.Caption = "Connected to server."
100           lblRetrievingUpdateInfo.visible = True
110           DoEvents
120       Else
130           DoEvents
140       End If

End Sub


Private Function ReadFileSize(ByRef http As Inet, ByVal url As String) As Long

          Dim strHeader As String
          
10        With http
20            .Protocol = icHTTP
30            .url = url
40            .Execute , "GET", , "Range: bytes=0-" & vbCrLf
          
50            While .StillExecuting
60                DoEvents
70            Wend
              
80        End With
          
90        strHeader = http.GetHeader("Content-Length")
          
100       http.Cancel
110       ReadFileSize = val(strHeader)
End Function

Private Sub HTTPDownloadFile(ByRef http As Inet, fileInfo As tFileToUpdate, totalsize As Long, doneSize As Long, tickStart As Long)
          
10        On Error GoTo HTTPDownloadFile_Error
          
          Const Chunk_Size As Long = 1024
          Dim f As Integer
          Dim strHeader As String
          Dim b() As Byte
          Dim tmpstr As String
          Dim fullstr As String
20        fullstr = ""
          Dim lngBytesReceived As Long
          Dim lngFileLength As Long
          
30        DoEvents
40        With http
50            .Protocol = icHTTP
60            .url = fileInfo.url
70            .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
          
80            While .StillExecuting
90                DoEvents
100           Wend
              
110       End With
          
120       lngFileLength = fileInfo.filesize
          
130       DoEvents
          
140       lngBytesReceived = 0
150       f = FreeFile()
          
          'create needed folder
160       CreateDir GetPathTo(fileInfo.localpath)
          Dim lasttick As Long
          Dim strprogress As String
          
170       If FileExists(fileInfo.localpath) Then Kill (fileInfo.localpath)
          
180       If IsTextFile(fileInfo.localpath) Then
190           Do
200               tmpstr = http.GetChunk(Chunk_Size, icString)
210               fullstr = fullstr & tmpstr
220               lngBytesReceived = lngBytesReceived + Len(tmpstr)
                  
230               strprogress = "Downloading " & GetFileTitle(fileInfo.localpath) & " " & IIf(fileInfo.filesize > 0, " (" & Int(lngBytesReceived / (fileInfo.filesize + 1) * 100) & "% of " & IIf(fileInfo.filesize > 1048576, Format(fileInfo.filesize / 1024 / 1024, "0.00") & "MB)", fileInfo.filesize \ 1024 & "KB)"), "")
240               If lblRetrievingUpdateInfo.Caption <> strprogress Then
250                   lblRetrievingUpdateInfo.Caption = strprogress
260               End If
                  
270               If totalsize > 0 Then
280                   progress.value = doneSize + lngBytesReceived
                      
290                   If GetTickCount - lasttick > 200 Or lngBytesReceived >= lngFileLength Then 'Update rate
300                       lblGlobalProgress.Caption = "Total Progress: " & Int(((doneSize + lngBytesReceived) / totalsize) * 100) & "% " & IIf(GetTickCount - tickStart > 0, " @ " & Format((doneSize + lngBytesReceived) / (GetTickCount - tickStart), "#") & "KB/s", "")
310                       lasttick = GetTickCount
320                   End If
                      
330               End If
340           Loop While Len(tmpstr) > 0
              
350           fullstr = replace(fullstr, Chr$(10), vbNewLine)
360           Open fileInfo.localpath For Output As #f
370               Print #f, fullstr
380           Close #f
390       Else
400           Open fileInfo.localpath For Binary Access Write As #f
410           Do
420               b = http.GetChunk(Chunk_Size, icByteArray)
430               Put #f, , b
440               lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1
                  
450               strprogress = "Downloading " & GetFileTitle(fileInfo.localpath) & " " & IIf(fileInfo.filesize > 0, " (" & Int(lngBytesReceived / (fileInfo.filesize + 1) * 100) & "% of " & IIf(fileInfo.filesize > 1048576, Format(fileInfo.filesize / 1024 / 1024, "0.00") & "MB)", fileInfo.filesize \ 1024 & "KB)"), "")
460               If lblRetrievingUpdateInfo.Caption <> strprogress Then
470                   lblRetrievingUpdateInfo.Caption = strprogress
480               End If
                  
490               If totalsize > 0 Then
500                   progress.value = doneSize + lngBytesReceived
                      
510                   If GetTickCount - lasttick > 200 Or lngBytesReceived >= lngFileLength Then 'Update rate
520                       lblGlobalProgress.Caption = "Total Progress: " & Int(((doneSize + lngBytesReceived) / totalsize) * 100) & "% " & IIf(GetTickCount - tickStart > 0, " @ " & Format((doneSize + lngBytesReceived) / (GetTickCount - tickStart), "#") & "KB/s", "")
530                       lasttick = GetTickCount
540                   End If
                      
550               End If
560           Loop While UBound(b, 1) > 0
570           Close #f
580       End If
          
          
590       lngFileLength = 0
600       lngBytesReceived = 0

610       On Error GoTo 0
620       Exit Sub
          
HTTPDownloadFile_Error:
630       HandleError Err, "frmCheckUpdate.HTTPDownloadFile", False
          
End Sub



Private Sub txtDetails_Change()
10        txtDetails.selstart = Len(txtDetails.Text)
End Sub
