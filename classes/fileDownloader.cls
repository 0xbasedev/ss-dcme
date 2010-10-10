VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fileDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim c_BytesDownloaded As Long
Dim c_BytesToDownload As Long
Dim c_LocalFileName As String
Dim c_FileURL As String

Private ctlInet As Object
Attribute ctlInet.VB_VarHelpID = -1

'InetCtlsObjects.Inet

Sub DownloadFile(url As String, localname As String)
    c_FileURL = url
    c_LocalFileName = localname
    
    
    
End Sub

Function GetDownloadedBytes() As Long
    GetDownloadedBytes = c_BytesDownloaded
End Function

Function GetFileSize() As Long
    GetFileSize = c_BytesToDownload
End Function

Function GetProgressPercent() As Double
    If c_BytesToDownload = 0 Then
        GetProgressPercent = 0#
    Else
        GetProgressPercent = c_BytesDownloaded / c_BytesToDownload
    End If
End Function

Private Sub Class_Initialize()
    c_BytesDownloaded = 0
    c_BytesToDownload = 0
    c_LocalFileName = vbNullString
    c_FileURL = vbNullString
    
End Sub


Sub InitObject(Optional protocol As ProtocolConstants = icHTTP)
    Set ctlInet = CreateObject("InetCtls.Inet")
    
    ctlInet.Cancel
    
    ctlInet.protocol = protocol
End Sub

Sub ClearObject()
    Set ctlInet = Nothing
End Sub
Function ReadHTTPTextFile(ByVal sURL As String) As String
    InitObject
    
    ' get the page
    ReadHTTPTextFile = ctlInet.OpenURL(sURL)
    
    ClearObject
End Function



Private Function ReadFileSize(ByRef http As Inet, ByVal url As String) As Long
    Dim strHeader As String
    
    With http
        .protocol = icHTTP
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
'    Const Chunk_Size As Long = 1024
'    Dim f As Integer
'    Dim strHeader As String
'    Dim b() As Byte
'    Dim tmpstr As String
'    Dim fullstr As String
'    fullstr = ""
'    Dim lngBytesReceived As Long
'    Dim lngFileLength As Long
'
'    DoEvents
'    With http
'        .protocol = icHTTP
'        .url = fileInfo.url
'        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
'
'        While .StillExecuting
'            DoEvents
'        Wend
'
'    End With
'
'    lngFileLength = fileInfo.filesize
'
'    DoEvents
'
'    lngBytesReceived = 0
'    f = FreeFile()
'
'    'create needed folder
'    CreateDir GetPathTo(fileInfo.localpath)
'    Dim lasttick As Long
'    Dim strprogress As String
'
'    If FileExists(fileInfo.localpath) Then Kill (fileInfo.localpath)
'
'    If IsTextFile(fileInfo.localpath) Then
'        Do
'            tmpstr = http.GetChunk(Chunk_Size, icString)
'            fullstr = fullstr & tmpstr
'            lngBytesReceived = lngBytesReceived + Len(tmpstr)
'
'            strprogress = "Downloading " & GetFileTitle(fileInfo.localpath) & " " & IIf(fileInfo.filesize > 0, " (" & Int(lngBytesReceived / (fileInfo.filesize + 1) * 100) & "% of " & IIf(fileInfo.filesize > 1048576, Format(fileInfo.filesize / 1024 / 1024, "0.00") & "MB)", fileInfo.filesize \ 1024 & "KB)"), "")
'            If lblRetrievingUpdateInfo.Caption <> strprogress Then
'                lblRetrievingUpdateInfo.Caption = strprogress
'            End If
'
'            If totalsize > 0 Then
'                ProgressBar1.value = Int(((doneSize + lngBytesReceived) / totalsize) * 100)
'
'                If GetTickCount - lasttick > 200 Or lngBytesReceived >= lngFileLength Then 'Update rate
'                    lblGlobalProgress.Caption = "Total Progress: " & Int(((doneSize + lngBytesReceived) / totalsize) * 100) & "% - " & Format((doneSize + lngBytesReceived) \ 1024, "# ###") & "KB / " & Format((totalsize \ 1024), "# ###") & "KB" & IIf(GetTickCount - tickStart > 0, " - @" & Format((doneSize + lngBytesReceived) / (GetTickCount - tickStart), "#") & "KB/s", "")
'                    lasttick = GetTickCount
'                End If
'
'            End If
'        Loop While Len(tmpstr) > 0
'
'        fullstr = replace(fullstr, Chr$(10), vbNewLine)
'        Open fileInfo.localpath For Output As #f
'            Print #f, fullstr
'        Close #f
'    Else
'        Open fileInfo.localpath For Binary Access Write As #f
'        Do
'            b = http.GetChunk(Chunk_Size, icByteArray)
'            Put #f, , b
'            lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1
'
'            strprogress = "Downloading " & GetFileTitle(fileInfo.localpath) & " " & IIf(fileInfo.filesize > 0, " (" & Int(lngBytesReceived / (fileInfo.filesize + 1) * 100) & "% of " & IIf(fileInfo.filesize > 1048576, Format(fileInfo.filesize / 1024 / 1024, "0.00") & "MB)", fileInfo.filesize \ 1024 & "KB)"), "")
'            If lblRetrievingUpdateInfo.Caption <> strprogress Then
'                lblRetrievingUpdateInfo.Caption = strprogress
'            End If
'
'            If totalsize > 0 Then
'                ProgressBar1.value = Int(((doneSize + lngBytesReceived) / totalsize) * 100)
'
'                If GetTickCount - lasttick > 200 Or lngBytesReceived >= lngFileLength Then 'Update rate
'                    lblGlobalProgress.Caption = "Total Progress: " & Int(((doneSize + lngBytesReceived) / totalsize) * 100) & "% - " & Format((doneSize + lngBytesReceived) \ 1024, "# ###") & "KB / " & Format((totalsize \ 1024), "# ###") & "KB" & IIf(GetTickCount - tickStart > 0, " - @" & Format((doneSize + lngBytesReceived) / (GetTickCount - tickStart), "#") & "KB/s", "")
'                    lasttick = GetTickCount
'                End If
'
'            End If
'        Loop While UBound(b, 1) > 0
'        Close #f
'    End If
'
'
'    lngFileLength = 0
'    lngBytesReceived = 0

End Sub
