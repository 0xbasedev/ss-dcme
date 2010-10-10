VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eLVLdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Metadata_Header
    magic(3) As Byte    '("elvl", for "extended lvl")
    totalsize(3) As Byte    'number of bytes of the whole metadata section
    Reserved As Long    'should be 0
End Type


Private Type Chunk_Header
    chunktype(3) As Byte    'chunk type
    size(3) As Byte    'number of bytes in the data portion of the chunk
    'NOT including the header
End Type


Private Type ATTR_type
    Key As String
    value As String
    ''    name As String
    ''    version As String
    ''    zone As String
    ''    mapcreator As String '"name <email>"
    ''    tilesetcreator As String
    ''    program As String
End Type




'ATTR
Dim Attributes() As ATTR_type
Public attrCount As Integer

'Unknown chunks
Dim Unknown() As Unknown_chunk
Public unknownCount As Integer

Dim parent As frmMain

Dim HashCode As String

Sub setParent(Main As frmMain)
    Set parent = Main
End Sub


''''''''''''''''''''''''''''''''''
''''''''ATTRIBUTES''''''''''''''''
''''''''''''''''''''''''''''''''''
Sub NewAttribute(Key As String, value As String)
    ReDim Preserve Attributes(attrCount)

    Attributes(attrCount).Key = Key
    Attributes(attrCount).value = value

    attrCount = attrCount + 1

End Sub

Function GetAttributeKey(Index As Integer) As String
    GetAttributeKey = Attributes(Index).Key
End Function

Function GetAttributeValue(Index As Integer) As String
    GetAttributeValue = Attributes(Index).value
End Function

Sub ClearAttributes()
    Erase Attributes
    attrCount = 0
End Sub



''''''''''''''''''''''''''''''''''
''''''''REGIONS'''''''''''''''''''
''''''''''''''''''''''''''''''''''


'Private Type REGN_rTIL_type
'    'explanation = wtf ?
'    b() As Byte
'End Type
'
'Private Type REGN_rAWP_type
'    X As Integer
'    Y As Integer
'    arena As String 'Max 16 bytes! IF SAME ARENA, DO NOT INCLUDE
'End Type
'
'Private Type REGN_type
'    name As String
'    tileInfo As REGN_rTIL_type '" tile data, the definition of the region"
'    base As Boolean 'IF NO BASE, DO NOT INCLUDE - whether the region represents a base in a flag game
'    antiwarp As Boolean 'IF ANTIWARP,DO NOT INCLUDE - no antiwarp
'    weapons As Boolean 'IF WEAPONS, DO NOT INCLUDE - no weapons
'    flagdrop As Boolean 'IF FLAG DROP, DO NOT INCLUDE - no flag drops
'    autowarp As REGN_rAWP_type '- auto-warp
'    PYcode As String ' code to be executed when a player enters or leaves this region
'End Type
'
''TSET and TILE not included
'
'Dim attr() As ATTR_type
'Dim regn() As REGN_type
'

Sub GetELVLData(f As Integer, metadata_position As Long, filename As String)
    'Grabs all the eLVL data from file f, starting at given position
    
    Dim Metadata_Header As Metadata_Header
    Dim Chunk_Header As Chunk_Header

    Dim CurrentPosition As Long
    Dim Metadata_Size As Long
    Dim Chunk_Type As String
    Dim Chunk_Size As Long
    Dim Chunk_Data() As Byte

    Dim Region As Region
    Dim regnCount As Integer
    Dim dcwtCount As Integer
    Dim i As Integer
    
    HashCode = ""
    
    ReDim Unknown(0) As Unknown_chunk
    unknownCount = 0
    
    Chunk_Size = 1

    'Start of eLVL data
    Seek #f, (metadata_position + 1)

    Get #f, , Metadata_Header


    Metadata_Size = bytesToLong(Metadata_Header.totalsize)

    parent.AddDebug "eLVL.GetELVLData, eLVL header info: " & vbCrLf & _
                    "                  --- magic: " & bytesToString(Metadata_Header.magic) & vbCrLf & _
                    "                  --- size: " & Metadata_Size & vbCrLf & _
                    "                  --- reserved: " & Metadata_Header.Reserved


    'Check if the magic number is correct
    If bytesToString(Metadata_Header.magic) = "elvl" Then
        'we have elvl data, proceed
        CurrentPosition = Seek(f)
        
        On Error GoTo InvalidELVLChunk
        
        While CurrentPosition < metadata_position + Metadata_Size And Chunk_Size > 0
            Get #f, , Chunk_Header

            Chunk_Type = bytesToString(Chunk_Header.chunktype)
            Chunk_Size = bytesToLong(Chunk_Header.size)
            
            If CurrentPosition + Chunk_Size - 8 > metadata_position + Metadata_Size Then
                MessageBox "Invalid eLVL chunk size", vbExclamation + vbOKOnly
                Chunk_Size = 0 'make the loop abort
                
            ElseIf Len(Chunk_Type) = 4 And Chunk_Size > 0 Then
                'chunk is most likely to be valid
                

                parent.AddDebug "eLVL.GetELVLData, Chunk loaded: " & Chunk_Type & " (" & Chunk_Size & ") at " & Seek(f) - 8

                If Chunk_Type = "ATTR" Then
                    Dim tmpString() As String
                    ReDim Chunk_Data(Chunk_Size - 1) As Byte

                    Get #f, , Chunk_Data

                    tmpString = Split(bytesToString(Chunk_Data), "=")

                    parent.AddDebug tmpString(0) & " = " & tmpString(1)
                    
                    Call NewAttribute(tmpString(0), tmpString(1))


                ElseIf Chunk_Type = "REGN" Then
                    'add new region
                    Set Region = New Region
                    
'                    ReDim Preserve Region(regnCount)  '(amountofregions - (amountofregions Mod 10) + 10)
'                    Set Region(regnCount) = New Region
                    
                    Call loadRegionInMap(f, Chunk_Size, Region)
                    'Region was loaded successfully
                    Call parent.Regions.setRegion(Region, regnCount)
                    regnCount = regnCount + 1

                    Set Region = Nothing

                ElseIf Chunk_Type = "DCWT" Then 'DCME walltiles data
                    If Chunk_Size = 16 Then 'Standard 16-tiles set
                    
                        ReDim Chunk_Data(Chunk_Size - 1) As Byte
                        Get #f, , Chunk_Data
                        
                        Call parent.walltiles.AddWalltileset(Chunk_Data, Chunk_Size)
                        
                    'elseif ... 'other walltiles formats
                    Else
                        Err.Raise 1, , "Unknown Walltiles set format: " & Chunk_Size
                    End If
                    
                    
                ElseIf Chunk_Type = "DCTT" Then 'DCME texttiles data
                    If Chunk_Size <= 512 Then
                        ReDim Chunk_Data(Chunk_Size - 1) As Byte
                        Get #f, , Chunk_Data
                        
                        For i = 0 To Chunk_Size - 1 Step 2
                            Call parent.TileText.SetalphaTile(CInt(Chunk_Data(i)), CInt(Chunk_Data(i + 1)))
                        Next
                    Else
                        Err.Raise 1, , "Invalid DCTT chunk size: " & Chunk_Size
                    End If
                    
                ElseIf Chunk_Type = "DCID" Then
                    'DCME hash code
                    ReDim Chunk_Data(Chunk_Size - 1) As Byte
                    
                    Get #f, , Chunk_Data
                    
                    HashCode = bytesToString(Chunk_Data)
                    
                    parent.AddDebug "MapID = " & HashCode
                    
                ElseIf Chunk_Type = "DCBM" Then
                    'Bookmarks
                    If Chunk_Size \ 4 <= 20 Then
                        ReDim Chunk_Data(3) As Byte
                        
                        For i = 1 To Chunk_Size \ 4
                            Get #f, , Chunk_Data 'retrieve 4 bytes
                        
                            Call parent.SetBookMark(CInt(Chunk_Data(3)), _
                                                    Chunk_Data(0) + 256 * (Chunk_Data(1) Mod 16), _
                                                    Chunk_Data(1) \ 16 + 16 * Chunk_Data(2))
                            'extract the data
                        Next
                    Else
                        Err.Raise 1, , "Invalid DCBM chunk size: " & Chunk_Size
                    End If
                    
                ElseIf Chunk_Type = "DCLV" And (Chunk_Size >= 33) Then
                    'LVZ path
                    
                    Dim md5 As New clsMD5
                    Dim MD5hash() As Long
                    ReDim MD5hash(3)
                    Get #f, , MD5hash
                    
                    ReDim Chunk_Data(Chunk_Size - 33) As Byte
                    Get #f, , Chunk_Data
                    
                    Dim lvzpath As String
                    
                    lvzpath = bytesToString(Chunk_Data)
                    
                    'this should be a relative path
                    Dim lvzfullpath As String
                    
                    If InStr(lvzpath, ":") Then
                        'absolute path
                        lvzfullpath = lvzpath
                    Else
                        'relative path
                        lvzfullpath = GetPathTo(filename) & lvzpath
                    End If
                    
                    
                    
                    If FileExists(lvzfullpath) Then
                        md5.CalculateMD5 (lvzfullpath)
                        parent.AddDebug "LevelFile = " & lvzfullpath & "(" & md5.MD5Value & " ; Expecting " & md5.GetMD5Value(MD5hash) & ")"
                        
                        If md5.CompareWith(MD5hash) = False Then
                            If MessageBox(lvzfullpath & " was modified outside of DCME. Do you still want to load it?", vbYesNo) = vbYes Then
                                Call parent.lvz.importLVZ(lvzfullpath)
                            End If
                        Else
                            Call parent.lvz.importLVZ(lvzfullpath)
                        End If
                    Else
                        parent.AddDebug "LevelFile = " & lvzfullpath & "(" & md5.GetMD5Value(MD5hash) & ") not found!"
                        
                        If MessageBox(lvzfullpath & " was not found. Do you want to search it manually?", vbYesNo + vbQuestion, "File not found") = vbYes Then
                            'manual search
                            On Error GoTo GetELVLData_Error
                            
                            frmGeneral.cd.filename = lvzfullpath
                            
                            frmGeneral.cd.DialogTitle = "Search for " & lvzpath
                            
                            frmGeneral.cd.flags = cdlOFNHideReadOnly
                            frmGeneral.cd.Filter = "LVZ files (*.lvz)|*.lvz|All files (*.*)|*.*"
                            frmGeneral.cd.ShowOpen
                            lvzfullpath = frmGeneral.cd.filename
                            
                            If FileExists(lvzfullpath) Then
                                Call parent.lvz.importLVZ(lvzfullpath)
                            End If
                            
                            On Error GoTo InvalidELVLChunk
                        End If
                    End If
                        
                Else
                    'Unknown chunk type, seek back to the header, and copy the whole thing
                    If unknownCount >= UBound(Unknown) Then
                        ReDim Preserve Unknown(unknownCount + 10) As Unknown_chunk
                    End If
                    
'                    ReDim Chunk_Data(Chunk_Size - 1) As Byte
                    ReDim Unknown(unknownCount).Data(Chunk_Size - 1) As Byte

                    'gather data of the unknown chunk, including the header so we don't have to bother with its type
                    Seek #f, CurrentPosition
                    Get #f, , Unknown(unknownCount).Data

                    Unknown(unknownCount).Type = Chunk_Type
                    Unknown(unknownCount).size = Chunk_Size
                    Unknown(unknownCount).Data = Chunk_Data
                    unknownCount = unknownCount + 1
                End If


                'Seek next chunk header
                Seek #f, Next4bytes(Seek(f))
                CurrentPosition = Seek(f)
            Else
                Chunk_Size = 0
                'invalid chunk, loop will abort with a chunk_size of 0
            End If
        Wend

        'load regions in map if we have any
'        If regnCount > 0 Then
'            Call parent.Regions.setRegions(Region, regnCount - 1)

'            Call parent.Regions.BuildRegionTiles

'        End If


    Else
        'this is not eLVL data, seek back to start
        Seek #f, (metadata_position + 1)
    End If

    'if no hashcode was found, generate one
    If HashCode = "" Then
        Call GenerateHashCode
    End If
    
    On Error GoTo 0
    Exit Sub

GetELVLData_Error:
    If Err = cdlCancel Then
        Resume Next
    Else
        GoTo InvalidELVLChunk
    End If
InvalidELVLChunk:
    HandleError Err, "eLVLdata.GetELVLData:Invalid eLVL chunk", False, False
End Sub





Function PutELVLData(f As Integer, Optional flags As saveFlags = SFdefault) As Long
    'Saves all eLVL-related data into file f
    'Returns the total length of the data put
    
    Dim b() As Byte
    Dim i As Integer, j As Integer
    
    Const magic = &H6C766C65    '1819700325

    Dim totalsizePOS As Long       'Where to rewind to write the total size once done
    Dim startPOS As Long           'Start position of eLVL data

    Dim totalsize As Long          'keeps track of total size of eLVL data

    Dim customObjectsCount As Long 'used to count occurences of various elements
    
    '------ metadata header--------------
    startPOS = Seek(f)
    Put #f, , magic

    'store pos of totalsize
    totalsizePOS = Seek(f)
    'continue writing (write incorrect totalsize for now, so the sequence is OK
    Put #f, , totalsize
    Put #f, , CLng(0)    'reserved space = 0
    '------ end metadata header ----------



    'ATTR CHUNKS
    If FlagIs(flags, SFsaveELVLattr) And attrCount > 0 Then

        'Dim attrSizePOS As Long
        'Dim attrSize As Long
        Dim attrString As String

        For i = 0 To attrCount - 1
            If Attributes(i).Key <> "" Then
                attrString = Attributes(i).Key & "=" & Attributes(i).value

                parent.AddDebug "eLVL.PutELVLData, Saving Attribute(" & i & ") at " & Seek(f)
                '------ ATTR header ----------------
                Put #f, , stringToBytes("ATTR")

                Put #f, , CLng(Len(attrString))    'ATTR chunk size
                '------ end ATTR header


                '------ ATTR body ------------------


                Put #f, , stringToBytes(attrString)

                Seek #f, Next4bytes(Seek(f))
            End If
            '------ end ATTR chunk -------------

        Next

    End If
    '---------------------------------


    'REGN CHUNKS
    If FlagIs(flags, SFsaveELVLregn) Then
        Call saveRegionsFromMap(f)
    End If


    'HASHCODE CHUNK
    '---header
    parent.AddDebug "eLVL.PutELVLData, Saving Hashcode " & HashCode & " at " & Seek(f)
    
    Put #f, , stringToBytes("DCID")
    Put #f, , CLng(Len(HashCode))
    
    Put #f, , stringToBytes(HashCode)
    
    Seek #f, Next4bytes(Seek(f))
    

    'BOOKMARKS
    customObjectsCount = 0
    
    If FlagIs(flags, SFsaveELVLdcbm) Then
        Dim customBookmarks(9) As Boolean
        For i = 0 To 9
            If parent.GetBookMarkX(i) <> parent.GetBookMarkDefaultX(i) Or _
                parent.GetBookMarkY(i) <> parent.GetBookMarkDefaultY(i) Then
                customBookmarks(i) = True
                customObjectsCount = customObjectsCount + 1
            End If
        Next
        
        If customObjectsCount > 0 Then
            'put header
            Put #f, , stringToBytes("DCBM")
            Put #f, , CLng(customObjectsCount * 4) '4 bytes per bookmark
            
            ReDim b(3) As Byte
            Dim X As Integer
            Dim Y As Integer
            For i = 0 To 9
                If customBookmarks(i) Then
                    parent.AddDebug "eLVL.PutELVLData, Saving DCBM(" & i & ") at " & Seek(f)
                    
                    X = parent.GetBookMarkX(i)
                    Y = parent.GetBookMarkY(i)
                    
                    'build same way as tile data
                    'x:12 , y:12 , bookmarkID:8
                    b(0) = X Mod 256    'bottom 8 of x
                    b(2) = Y \ 16    'top 8 of y
                    b(1) = (Y - 16 * b(2)) * 16 + (X - b(0)) \ 256
                    b(3) = CByte(i)
                    
                    'now put the bytes
                    Put #f, , b(0)
                    Put #f, , b(1)
                    Put #f, , b(2)
                    Put #f, , b(3)
                End If
            Next
            
            Erase b
            
            Seek #f, Next4bytes(Seek(f))
        End If
    End If
        
    'DCWT Walltiles chunks
    
    If FlagIs(flags, SFsaveELVLdcwt) Then
    
        'TODO: if the walltile sets number becomes dynamic, this will need to be changed
        For i = 0 To 7
            If parent.walltiles.isValidSet(i) Then
                parent.AddDebug "eLVL.PutELVLData, Saving DCWT(" & i & ") at " & Seek(f)
                
                Put #f, , stringToBytes("DCWT")
                Put #f, , CLng(16)
                
                For j = 0 To 15
                    Put #f, , CByte(parent.walltiles.getWallTile(i, j))
                Next
                
                Seek #f, Next4bytes(Seek(f))
            End If
        Next
    End If
    
    'DCTT Texttile chunks
    If FlagIs(flags, SFsaveELVLdctt) Then
    
        ReDim b(511) As Byte
        customObjectsCount = 0
        
        For i = 0 To 255
            If parent.TileText.isCharacterCustom(i) Then
                b(customObjectsCount) = CByte(i)
                b(customObjectsCount + 1) = CByte(parent.TileText.GetalphaTile(i))
                customObjectsCount = customObjectsCount + 2
            End If
        Next
        
        If customObjectsCount > 0 Then
            parent.AddDebug "eLVL.PutELVLData, Saving DCTT at " & Seek(f)
            
            ReDim Preserve b(customObjectsCount - 1)
            Put #f, , stringToBytes("DCTT")
            Put #f, , customObjectsCount
            Put #f, , b
            
            Seek #f, Next4bytes(Seek(f))
        End If
        
        Erase b
    
    End If
    
    
    'DCLV lvz paths chunks
    If FlagIs(flags, SFsaveELVLdclv) Then
        For i = 0 To parent.lvz.getLVZCount - 1
            Dim lvzname As String
            lvzname = parent.lvz.getLVZname(i)
            parent.AddDebug "eLVL.PutELVLData, Saving LVZ '" & lvzname & "' at " & Seek(f)
            
            Put #f, , stringToBytes("DCLV")
            Put #f, , CLng(32 + Len(lvzname))
            
            'MD5 hash
            Dim md5 As New clsMD5
            Put #f, , md5.CalculateMD5(parent.lvz.getLVZpath(i))
            
            'Lvz name (relative path)
            Put #f, , stringToBytes(lvzname)
            
            Seek #f, Next4bytes(Seek(f))
        Next
    End If
    
    'UNKNOWN CHUNKS
    If FlagIs(flags, SFsaveELVLunknown) Then
        For i = 0 To unknownCount - 1
            parent.AddDebug "eLVL.PutELVLData, Saving Unknown(" & i & ") at " & Seek(f)
            
            Put #f, , stringToBytes(Unknown(i).Type)
            Put #f, , Unknown(i).size
            Put #f, , Unknown(i).Data
            
            Seek #f, Next4bytes(Seek(f))
        Next
    End If
    
    'Wtf is this? -Samapico
    Dim lastbyte As Byte
    Get #f, Seek(f) - 1, lastbyte
    If lastbyte = 0 Then
        Put #f, Seek(f) - 1, CByte(0)
    End If



    ' end of ELVL, seek back and overwrite size
    Dim curSeek As Long
    curSeek = Seek(f)

    'this is needed, else it crashes with empty maps because eLVL data
    'Put #f, , CByte(0)

    'update total size
    Seek #f, totalsizePOS
    totalsize = curSeek - startPOS

    parent.AddDebug "eLVL.PutELVLData, Rewriting total size: " & totalsize & " at " & Seek(f)
    Put #f, , CLng(LongToUnsigned(totalsize))



    'forward back to current pos
    Seek #f, curSeek

    'return totalsize
    PutELVLData = totalsize
End Function



Private Sub Class_Initialize()
    attrCount = 0
    unknownCount = 0
    
    'Generate a new hashcode for the map
    'If the map is being loaded, this hashcode will be replaced by the one loaded from eLVL data
    Call GenerateHashCode
End Sub



'INTERFACE SUBS/FUNCTIONS NEEDED BY THE REST OF THE PROGRAM

Sub getAttributeList(lstItems As ListItems)
'    'puts all the properties in the given list
'
    Dim i As Integer

    If attrCount > 0 Then
        For i = 0 To attrCount - 1
            '        ' RETRIEVE STUFF HERE
            '
            '
            Call lstItems.add(i + 1, Attributes(i).Key, Attributes(i).Key)

            lstItems.item(Attributes(i).Key).SubItems(1) = Attributes(i).value
        Next
    End If
End Sub

Sub setAttributeList(lstItems As ListItems)
'    'puts all the properties in eLVL data
    Call ClearAttributes

    Dim i As Integer
    '
    '    'first element in listview is always 1->count not 0->count-1 !!!!
    For i = 1 To lstItems.count
        Call NewAttribute(lstItems.item(i).Text, lstItems.item(i).SubItems(1))
    Next

End Sub

'Drake:
'Regions  will be stored as map regions in the map itself
'so it isn't necessary to keep the original chunks in memory
'they will be overwritten anyway when all regions from map are
'saved back into the chunks
'Sama:
'Only the unknown-type chunks are saved in memory, other
'chunks are converted to usable data in GeteLVLdata

Private Sub loadRegionInMap(f As Integer, totalsize As Long, ByRef tmpRegion As Region)
'    'converts the regions from elvl to map regions (saved in the map itself) so
'    'it hasn't have to be converted each time we need a region (would slow down too
'    'much
'
'    'this will be called at LoadMap
'
    Dim CurrentPosition As Long    'Current position in file
    Dim StartPosition As Long

    
    'Sub-chunk information
    Dim SubChunk_Header As Chunk_Header
    Dim SubChunk_Type As String
    Dim SubChunk_Size As Long
    Dim SubChunk_Data() As Byte
    
    
    Dim regnUnknownChunks() As Unknown_chunk
    Dim regnUnknownCount As Long
    
    regnUnknownCount = 0
    
'    Call loadRegionInMap.bitfield.InitArray(1024, 1024, 0)
'    'Dim defaultbitfield(1023, 1023) As Boolean
'    'loadRegionInMap.bitfield = defaultbitfield
'
    Randomize
'
'    'size of regions is always a multiplication of 10
'    'so 8 regions --> 10, 11 regions --> 20, 15 regions --> 20 etc
'    '????????
'    '????????
'
'    tmpRegion.unknownCount = 0
'    ReDim tmpRegion.unknownChunk(0)
    
    Call tmpRegion.InitRegion("", RGB(50 + Rnd() * 200, 50 + Rnd() * 200, 50 + Rnd() * 200))
    
'    tmpRegion.color = RGB(50 + Rnd() * 200, 50 + Rnd() * 200, 50 + Rnd() * 200)
'    tmpRegion.visible = True

    CurrentPosition = Seek(f)
    StartPosition = Seek(f)
    'read subchunks until end of REGN chunk
    While CurrentPosition < StartPosition + totalsize    'And SubChunk_Size > 0 'some subchunks might have a size of 0

        Get #f, , SubChunk_Header

        SubChunk_Type = bytesToString(SubChunk_Header.chunktype)
        SubChunk_Size = bytesToLong(SubChunk_Header.size)

        parent.AddDebug "   eLVL.loadRegionInMap, SubChunk loaded: " & SubChunk_Type & " (" & SubChunk_Size & ") at " & Seek(f) - 8

        If SubChunk_Size > 0 Then

            'Region name chunk
            If SubChunk_Type = "rNAM" Then
                ReDim SubChunk_Data(SubChunk_Size - 1)
                Get #f, , SubChunk_Data

                tmpRegion.name = bytesToString(SubChunk_Data)

                'Tile data chunk
            ElseIf SubChunk_Type = "rTIL" Then
                ReDim SubChunk_Data(SubChunk_Size - 1)
                Get #f, , SubChunk_Data

                'loadRegionInMap.bitfield = LoadRegionTileData(SubChunk_Data)
                Call LoadRegionTileData(SubChunk_Data, tmpRegion)

                'Python code
            ElseIf SubChunk_Type = "rPYC" Then
                ReDim SubChunk_Data(SubChunk_Size - 1)
                Get #f, , SubChunk_Data

                tmpRegion.pythonCode = bytesToString(SubChunk_Data)


                'Autowarp
            ElseIf SubChunk_Type = "rAWP" Then
                Dim warpX As Integer
                Dim warpY As Integer
                Get #f, , warpX
                Get #f, , warpY
                tmpRegion.autowarpX = warpX
                tmpRegion.autowarpY = warpY
                tmpRegion.isAutoWarp = True
                
                'Check for an arena name
                If SubChunk_Size > 4 Then
                    ReDim SubChunk_Data(SubChunk_Size - 5)
                    Get #f, , SubChunk_Data

                    tmpRegion.autowarpArena = bytesToString(SubChunk_Data)


                End If

                'Color
            ElseIf SubChunk_Type = "rCOL" Then
                ReDim SubChunk_Data(SubChunk_Size - 1)
                Get #f, , SubChunk_Data

                If UBound(SubChunk_Data) >= 2 Then
                    tmpRegion.color = RGB(CInt(SubChunk_Data(0)), CInt(SubChunk_Data(1)), CInt(SubChunk_Data(2)))
                End If

            Else
                'unknown subchunk type, store
                If regnUnknownCount = 0 Or regnUnknownCount >= UBound(regnUnknownChunks) Then
                    ReDim Preserve regnUnknownChunks(regnUnknownCount + 4)
                End If
                
'                    ReDim Chunk_Data(Chunk_Size - 1) As Byte
                ReDim regnUnknownChunks(regnUnknownCount).Data(SubChunk_Size - 1) As Byte
                
                Get #f, , regnUnknownChunks(regnUnknownCount).Data

                regnUnknownChunks(regnUnknownCount).Type = SubChunk_Type
                regnUnknownChunks(regnUnknownCount).size = SubChunk_Size

                regnUnknownCount = regnUnknownCount + 1
                
'                Seek #f, (Seek(f) + SubChunk_Size)
            End If

            Seek #f, Next4bytes(Seek(f))    'search next subchunk

            'these are valid 0-byte chunks
        ElseIf SubChunk_Type = "rBSE" Then
            tmpRegion.isBase = True

        ElseIf SubChunk_Type = "rNAW" Then
            tmpRegion.isNoAntiwarp = True

        ElseIf SubChunk_Type = "rNWP" Then
            tmpRegion.isNoWeapon = True

        ElseIf SubChunk_Type = "rNFL" Then
            tmpRegion.isNoFlagDrop = True



        Else
            'chunk is an unknown chunk
                'unknown subchunk type, store
            If regnUnknownCount = 0 Or regnUnknownCount >= UBound(regnUnknownChunks) Then
                ReDim Preserve regnUnknownChunks(regnUnknownCount + 4)
            End If
            
            'This chunk has no data
            
            regnUnknownChunks(regnUnknownCount).Type = SubChunk_Type
            regnUnknownChunks(regnUnknownCount).size = 0

            regnUnknownCount = regnUnknownCount + 1
            
        End If

        CurrentPosition = Seek(f)
        '"rNAM" - a descriptive name for the region
        '"rBSE" - whether the region represents a base in a flag game
        '"rNAW" - no antiwarp
        '"rNWP" - no weapons
        '"rNFL" - no flag drops
        '"rAWP" - auto-warp
        '"rPYC" - code to be executed when a player enters or leaves this region
        '"rTIL" - tile data, the definition of the region
    Wend
    
    If regnUnknownCount > 0 Then
    ReDim Preserve regnUnknownChunks(regnUnknownCount - 1)
        Call tmpRegion.SetUnknownChunks(regnUnknownChunks(), regnUnknownCount)
    End If
End Sub

Private Sub saveRegionsFromMap(f As Integer)
'saves all regions from the map to the chunks
'this will be called on SaveMap

'    Dim tmpRegion As Region
    
    Dim tick As Long
    tick = GetTickCount
    

    Dim nrOfRegions As Integer
    nrOfRegions = parent.Regions.getRegionIndex + 1

    Dim regionsize As Long
    Dim startPOS As Long

    

    Dim i As Integer


    For i = 0 To nrOfRegions - 1
        
'        Set tmpRegion = parent.Regions.getRegion(i)
'        Regions = parent.Regions.getRegionRegions
        
        parent.AddDebug "   eLVL.saveRegionsFromMap, Saving Region(" & i & ") at " & Seek(f)

        'PUT REGN HEADER

        Put #f, , stringToBytes("REGN")
        Put #f, , regionsize    'temporary size

        startPOS = Seek(f)


        Call parent.Regions.getRegion(i).SaveToFile(f)
        
        'WE FINISHED PUTTING DATA, UPDATE HEADER

        Dim curSeek As Long
        curSeek = Seek(f)

        'calculate totalsize
        regionsize = curSeek - startPOS

        'seek back to the size byte and overwrite it
        Seek #f, startPOS - 4
        Put #f, , CLng(LongToUnsigned(regionsize))

        'seek forward again to the end of the current region data
        Seek #f, Next4bytes(curSeek)
        parent.AddDebug "      eLVL.saveRegionsFromMap, End of region at " & Seek(f) & " - size was " & regionsize & " time: " & GetTickCount - tick

'        Set tmpRegion = Nothing
    Next

    
End Sub



Private Sub LoadRegionTileData(ByRef Data() As Byte, ByRef tmpRegion As Region)


    'Dim b(1023, 1023) As Boolean

    Dim tilesInRow As Integer
    Dim rowsCounted As Integer

    Dim i As Integer
    Dim j As Integer

    Dim CurrentPosition As Long
    Dim size As Long

    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim value As Integer

    'size of the rTIL data
    size = UBound(Data)

    While CurrentPosition <= size And rowsCounted < 1024

        Byte1 = Data(CurrentPosition)

        'first check the first 3 bits of the byte
        If Byte1 \ 32 = 0 Then
            '    000nnnnn          - n+1 (1-32) empty tiles in a row
            value = Byte1 Mod 32 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " false tiles @ (" & tilesInRow & "," & rowsCounted & ")"

'            All values are set to false by default already, no need to set these

'            If tilesInRow + Value > 1024 Then
'                'something's wrong, more than 1024 tiles in that row
'            Else
'                For i = tilesInRow To tilesInRow + Value - 1
'                    b(i, rowsCounted) = False
'                Next
'            End If

            tilesInRow = tilesInRow + value

            CurrentPosition = CurrentPosition + 1

        ElseIf Byte1 \ 32 = 1 Then
            '    001000nn nnnnnnnn - n+1 (1-1024) empty tiles in a row
            Byte2 = Data(CurrentPosition + 1)
            value = 256 * (Byte1 Mod 4) + Byte2 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " false tiles @ (" & tilesInRow & "," & rowsCounted & ")"

'            All values are set to false by default already, no need to set these

'            If tilesInRow + Value > 1024 Then
'                'something's wrong, more than 1024 tiles in that row
'            Else
'                For i = tilesInRow To tilesInRow + Value - 1
'                    b(i, rowsCounted) = False
'                Next
'            End If

            tilesInRow = tilesInRow + value

            CurrentPosition = CurrentPosition + 2

        ElseIf Byte1 \ 32 = 2 Then
            '    010nnnnn          - n+1 (1-32) present tiles in a row
            value = Byte1 Mod 32 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " true tiles @ (" & tilesInRow & "," & rowsCounted & ")"

            If tilesInRow + value > 1024 Then
                'something's wrong, more than 1024 tiles in that row
            Else
'                For i = tilesInRow To tilesInRow + Value - 1
'                    b(i, rowsCounted) = True
'                Next
'                Call bitfield.SetRectangle(tilesInRow, tilesInRow + Value - 1, rowsCounted, rowsCounted, 1)
                
                Call tmpRegion.AddRectangle(tilesInRow, rowsCounted, tilesInRow + value - 1, rowsCounted, Nothing, 0)
            End If

            tilesInRow = tilesInRow + value

            CurrentPosition = CurrentPosition + 1

        ElseIf Byte1 \ 32 = 3 Then
            '    011000nn nnnnnnnn - n+1 (1-1024) present tiles in a row
            Byte2 = Data(CurrentPosition + 1)
            value = 256 * (Byte1 Mod 4) + Byte2 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " true tiles @ (" & tilesInRow & "," & rowsCounted & ")"

            If tilesInRow + value > 1024 Then
                'something's wrong, more than 1024 tiles in that row
            Else
'                For i = tilesInRow To tilesInRow + Value - 1
'                    b(i, rowsCounted) = True
'                Next
'                Call bitfield.SetRectangle(tilesInRow, tilesInRow + Value - 1, rowsCounted, rowsCounted, 1)
                
                Call tmpRegion.AddRectangle(tilesInRow, rowsCounted, tilesInRow + value - 1, rowsCounted, Nothing, 0)
            End If

            tilesInRow = tilesInRow + value

            CurrentPosition = CurrentPosition + 2

        ElseIf Byte1 \ 32 = 4 Then
            '    100nnnnn          - n+1 (1-32) rows of all empty
            value = Byte1 Mod 32 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " empty rows @ (" & tilesInRow & "," & rowsCounted & ")"

            rowsCounted = rowsCounted + value

            CurrentPosition = CurrentPosition + 1

        ElseIf Byte1 \ 32 = 5 Then
            '    101000nn nnnnnnnn - n+1 (1-1024) rows of all empty
            Byte2 = Data(CurrentPosition + 1)
            value = 256 * (Byte1 Mod 4) + Byte2 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " empty rows @ (" & tilesInRow & "," & rowsCounted & ")"

            rowsCounted = rowsCounted + value

            CurrentPosition = CurrentPosition + 2

        ElseIf Byte1 \ 32 = 6 Then
            '    110nnnnn          - repeat last row n+1 (1-32) times
            value = Byte1 Mod 32 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " repeated rows @ (" & tilesInRow & "," & rowsCounted & ")"

'            For i = 0 To 1023
'                For j = rowsCounted To rowsCounted + Value - 1
'                    b(i, j) = b(i, j - 1)
'                Next
'            Next
            'Copy the entire row
            For i = 1 To value
'420                   Call tmpRegion.bitfield.MoveRectangle(0, 1023, rowsCounted - 1, rowsCounted - 1, 0, i, False)
                
                Call tmpRegion.CopyRow(rowsCounted - 1 + i - 1, rowsCounted - 1 + i)
                
'430                   BitBlt tmpRegion.disphDC, 0, rowsCounted - 1 + i, 1024, 1, tmpRegion.disphDC, 0, rowsCounted - 1, vbSrcCopy
            Next
            

            rowsCounted = rowsCounted + value

            CurrentPosition = CurrentPosition + 1

        ElseIf Byte1 \ 32 = 7 Then
            '    111000nn nnnnnnnn - repeat last row n+1 (1-1024) times
            Byte2 = Data(CurrentPosition + 1)
            value = 256 * (Byte1 Mod 4) + Byte2 + 1

'            parent.AddDebug "      eLVL.LoadRegionTileData, " & Value & " repeated rows @ (" & tilesInRow & "," & rowsCounted & ")"

'            For i = 0 To 1023
'                For j = rowsCounted To rowsCounted + Value - 1
'                    b(i, j) = b(i, j - 1)
'                Next
'            Next
            'Copy the entire row
            For i = 1 To value
'510                   Call tmpRegion.bitfield.MoveRectangle(0, 1023, rowsCounted - 1, rowsCounted - 1, 0, i, False)
                
                Call tmpRegion.CopyRow(rowsCounted - 1 + i - 1, rowsCounted - 1 + i)
            Next
            rowsCounted = rowsCounted + value

            CurrentPosition = CurrentPosition + 2

        End If

        If tilesInRow = 1024 Then
            tilesInRow = 0
            rowsCounted = rowsCounted + 1
        End If

    Wend


'    LoadRegionTileData = b
End Sub






'Does the map has any eLVL data?
Function HasData() As Boolean
    'Because of hash code, as soon as there is a tileset, consider that the map needs eLVL data
    HasData = Not parent.usingDefaultTileset Or _
              unknownCount > 0 Or attrCount > 0 Or parent.Regions.getRegionIndex >= 0
'        HasData = True
'    Else
'        HasData = False
'    End If
End Function






Function GenerateHashCode() As String
    Dim No&, S As String * MAX_FILENAME_LEN
    Call GetVolumeInformation("C:\", S, MAX_FILENAME_LEN, _
    No, 0&, 0&, S, MAX_FILENAME_LEN)

    GenerateHashCode = Hex$(No Mod 65521) & Hex$(CLng(Format(Now, "MMDDHHMMSS")) Mod 65521)
    
    HashCode = GenerateHashCode
End Function

Function GetHashCode() As String
    GetHashCode = HashCode
End Function







''''''''''''''''''''''''''''


'"rNAM" - a descriptive name for the region
'"rBSE" - whether the region represents a base in a flag game
'"rNAW" - no antiwarp
'"rNWP" - no weapons
'"rNFL" - no flag drops
'"rAWP" - auto-warp
'"rPYC" - code to be executed when a player enters or leaves this region
'"rTIL" - tile data, the definition of the region
'    000nnnnn          - n+1 (1-32) empty tiles in a row
'    001000nn nnnnnnnn - n+1 (1-1024) empty tiles in a row
'    010nnnnn          - n+1 (1-32) present tiles in a row
'    011000nn nnnnnnnn - n+1 (1-1024) present tiles in a row
'    100nnnnn          - n+1 (1-32) rows of all empty
'    101000nn nnnnnnnn - n+1 (1-1024) rows of all empty
'    110nnnnn          - repeat last row n+1 (1-32) times
'    111000nn nnnnnnnn - repeat last row n+1 (1-1024) times
'"rCOL" - DCME-custom ; color of region. Size 4. Format: (R,G,B,0)


'###########
'DCME CUSTOM CHUNKS
'###########
'DCID - DCME hash code, 8 chars
'DCWT - DCME wall tiles


'"rCOL" - DCME-custom ; color of region. Size 4. Format: (R,G,B,0)
Private Sub Class_Terminate()
    Erase Attributes
    Erase Unknown
    
    Set parent = Nothing
End Sub
