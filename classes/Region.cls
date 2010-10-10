VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Region"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim m_Flags As Long

Const F_ISBASE = &H1
Const F_ISNOANTIWARP = &H2
Const F_ISNOWEAPON = &H4
Const F_ISNOFLAGDROP = &H8
Const F_ISAUTOWARP = &H10

'Private m_objbitfield As boolArray
Private m_sname As String

Private m_iautowarpX As Integer
Private m_iautowarpY As Integer
Private m_sautowarpArena As String

Private m_spythonCode As String
Private m_lcolor As Long


Private Boundaries As area


Private disp As clsDisplayLayer
Public disphDC As Long

Private m_bvisible As Boolean

Dim unknownChunk() As Unknown_chunk
Private m_lunknownCount As Long



Private Sub Class_Initialize()
'    Set bitfield = New boolArray
'    Call bitfield.InitArray(1024, 1024, 0)
    
    Set disp = New clsDisplayLayer
    
    'Make monochrome
    Call disp.SetBitCount(1)
    disp.TextColor = m_lcolor
    disp.BackColor = vbBlack
    
    Call disp.Resize(1024, 1024, False)
    
    disphDC = disp.hDC

End Sub

Public Sub InitRegion(nName As String, nColor As Long)
    m_bvisible = True
    m_lcolor = nColor
    m_sname = nName
    
End Sub

Private Sub Class_Terminate()
    
    Erase unknownChunk
    

    Set disp = Nothing



'    Set m_objbitfield = Nothing
    
End Sub



Sub CopyRow(fromRow As Integer, toRow As Integer)

    BitBlt disp.hDC, 0, toRow, 1024, 1, disp.hDC, 0, fromRow, vbSrcCopy
   
End Sub


'Friend Property Get bitfield() As boolArray
'
'    Set bitfield = m_objbitfield
'
'End Property
'
'Friend Property Set bitfield(objbitfield As boolArray)
'
'    Set m_objbitfield = objbitfield
'
'End Property

Public Property Get name() As String

    name = m_sname

End Property

Public Property Let name(ByVal sname As String)

    m_sname = sname

End Property



Public Property Get isBase() As Boolean
    isBase = ((m_Flags And F_ISBASE) = F_ISBASE)
End Property

Public Property Let isBase(ByVal bisBase As Boolean)
    If bisBase Then
        m_Flags = (m_Flags Or F_ISBASE)
    Else
        m_Flags = (m_Flags And Not F_ISBASE)
    End If
End Property

Public Property Get isNoWeapon() As Boolean
    isNoWeapon = ((m_Flags And F_ISNOWEAPON) = F_ISNOWEAPON)
End Property

Public Property Let isNoWeapon(ByVal bisNoWeapon As Boolean)
    If bisNoWeapon Then
        m_Flags = (m_Flags Or F_ISNOWEAPON)
    Else
        m_Flags = (m_Flags And Not F_ISNOWEAPON)
    End If
End Property

Public Property Get isNoAntiwarp() As Boolean
    isNoAntiwarp = ((m_Flags And F_ISNOANTIWARP) = F_ISNOANTIWARP)
End Property

Public Property Let isNoAntiwarp(ByVal bisNoAntiwarp As Boolean)
    If bisNoAntiwarp Then
        m_Flags = (m_Flags Or F_ISNOANTIWARP)
    Else
        m_Flags = (m_Flags And Not F_ISNOANTIWARP)
    End If
End Property

Public Property Get isNoFlagDrop() As Boolean
    isNoFlagDrop = ((m_Flags And F_ISNOFLAGDROP) = F_ISNOFLAGDROP)
End Property

Public Property Let isNoFlagDrop(ByVal bisNoFlagDrop As Boolean)
    If bisNoFlagDrop Then
        m_Flags = (m_Flags Or F_ISNOFLAGDROP)
    Else
        m_Flags = (m_Flags And Not F_ISNOFLAGDROP)
    End If
End Property

Public Property Get isAutoWarp() As Boolean
    isAutoWarp = ((m_Flags And F_ISAUTOWARP) = F_ISAUTOWARP)
End Property

Public Property Let isAutoWarp(ByVal bisAutoWarp As Boolean)
    If bisAutoWarp Then
        m_Flags = (m_Flags Or F_ISAUTOWARP)
    Else
        m_Flags = (m_Flags And Not F_ISAUTOWARP)
    End If
End Property

Public Property Get color() As Long

    color = m_lcolor

End Property

Public Property Let color(ByVal lcolor As Long)
    'We have to redraw what we already have with that new color
    
    Call disp.ColorReplace(m_lcolor, lcolor)
    
    m_lcolor = lcolor

    disphDC = disp.hDC
End Property

Public Property Get pythonCode() As String

    pythonCode = m_spythonCode

End Property

Public Property Let pythonCode(ByVal spythonCode As String)

    m_spythonCode = spythonCode

End Property

Public Property Get autowarpArena() As String

    autowarpArena = m_sautowarpArena

End Property

Public Property Let autowarpArena(ByVal sautowarpArena As String)

    m_sautowarpArena = sautowarpArena

End Property

Public Property Get autowarpY() As Integer

    autowarpY = m_iautowarpY

End Property

Public Property Let autowarpY(ByVal iautowarpY As Integer)

    m_iautowarpY = iautowarpY

End Property

Public Property Get autowarpX() As Integer

    autowarpX = m_iautowarpX

End Property

Public Property Let autowarpX(ByVal iautowarpX As Integer)

    m_iautowarpX = iautowarpX

End Property

Public Property Get unknownCount() As Long

    unknownCount = m_lunknownCount

End Property

Public Property Let unknownCount(ByVal lunknownCount As Long)

    m_lunknownCount = lunknownCount

End Property


Friend Function GetUnknownChunks() As Unknown_chunk()
    GetUnknownChunks = unknownChunk
End Function

Friend Sub SetUnknownChunks(chunks() As Unknown_chunk, count As Long)
    unknownChunk = chunks
    m_lunknownCount = count
End Sub


Public Property Get visible() As Boolean

    visible = m_bvisible

End Property

Public Property Let visible(ByVal bvisible As Boolean)

    m_bvisible = bvisible

End Property


Public Property Get value(x As Integer, y As Integer) As Boolean
'    value = bitfield.value(X, Y)

    value = GetPixel(disphDC, x, y)
    
End Property

Public Property Let value(x As Integer, y As Integer, val As Boolean)
'    m_objbitfield.value(X, Y) = val
    
    Call SetPixel(disphDC, x, y, val)
    

End Property




'''''''''''END OF PROPERTIES

Public Sub AddTile(ByVal x As Integer, ByVal y As Integer, ByRef undoch As Changes, ByVal regionID As Integer)
    If Not undoch Is Nothing Then

        If Not value(x, y) Then Call undoch.AddTileChange(RegionAdd, x, y, regionID)
    
    End If
    

    Call SetPixel(disphDC, x, y, vbWhite)
End Sub

Public Sub RemoveTile(ByVal x As Integer, ByVal y As Integer, ByRef undoch As Changes, ByVal regionID As Integer)
    If Not undoch Is Nothing Then

        If value(x, y) Then Call undoch.AddTileChange(RegionRemove, x, y, regionID)
    
    End If
    

    Call SetPixel(disphDC, x, y, vbBlack)
End Sub


Public Sub AddRectangle(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByRef undoch As Changes, ByVal regionID As Integer)
    
   
    
    Dim i As Integer, j As Integer

    
    If Not undoch Is Nothing Then
        
        For j = y1 To y2
            For i = x1 To x2
                If Not value(i, j) Then Call undoch.AddTileChange(RegionAdd, i, j, regionID)
            Next
        Next
    
    End If
    
'    Call m_objbitfield.SetRectangle(x1, x2, y1, y2, 1)
    
    Call DrawFilledRectangle(disphDC, x1, y1, x2 + 1, y2 + 1, vbWhite)

End Sub

Public Sub RemoveRectangle(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByRef undoch As Changes, ByVal regionID As Integer)
    
    Dim i As Integer, j As Integer

    If Not undoch Is Nothing Then
        
        For j = y1 To y2
            For i = x1 To x2
                If value(i, j) Then Call undoch.AddTileChange(RegionRemove, i, j, regionID)
            Next
        Next
    
    End If
    
'    Call m_objbitfield.SetRectangle(x1, x2, y1, y2, 0)
    
    Call DrawFilledRectangle(disphDC, x1, y1, x2 + 1, y2 + 1, vbBlack)
    
End Sub





Public Function SaveToFile(f As Integer) As Long
    Dim chunkdata() As Byte
    Dim chunksize As Long
    Dim i As Integer

    
    'Name subchunk
    AddDebug "      Region.SaveToFile, Putting Name at " & Seek(f)

    Put #f, , stringToBytes("rNAM")
    Put #f, , CLng(Len(name))

    Put #f, , stringToBytes(name)
    Seek #f, Next4bytes(Seek(f))

    'Tile data subchunk
    AddDebug "      Region.SaveToFile, Putting Tiledata at " & Seek(f)

    Put #f, , stringToBytes("rTIL")

    chunksize = MakeTileData(chunkdata)

    Put #f, , chunksize

    Put #f, , chunkdata

    Seek #f, Next4bytes(Seek(f))


    'Autowarp subchunks
    If isAutoWarp Then
        AddDebug "      Region.SaveToFile, Putting Autowarp at " & Seek(f)

        Put #f, , stringToBytes("rAWP")

        If autowarpArena <> "" Then
            Put #f, , CLng(20)    '2 for X, 2 for Y, 16 for arena name
        Else
            Put #f, , CLng(4)
        End If

        Put #f, , autowarpX
        Put #f, , autowarpY

        If autowarpArena <> "" Then
            Put #f, , stringToBytes(Mid(autowarpArena, 1, 16))
            Seek #f, Seek(f) + 16 - Len(autowarpArena)
        End If

        Seek #f, Next4bytes(Seek(f))

    End If

    'Is Base
    If isBase Then
        AddDebug "      Region.SaveToFile, Putting IsBase at " & Seek(f)

        Put #f, , stringToBytes("rBSE")
        Put #f, , CLng(0)
    End If

    'No Antiwarp
    If isNoAntiwarp Then
        AddDebug "      Region.SaveToFile, Putting No AW at " & Seek(f)

        Put #f, , stringToBytes("rNAW")
        Put #f, , CLng(0)
    End If

    'No Weapons
    If isNoWeapon Then
        AddDebug "      Region.SaveToFile, Putting No Weapons at " & Seek(f)

        Put #f, , stringToBytes("rNWP")
        Put #f, , CLng(0)
    End If

    'No Flag drops
    If isNoFlagDrop Then
        AddDebug "      Region.SaveToFile, Putting No Flag Drops at " & Seek(f)

        Put #f, , stringToBytes("rNFL")
        Put #f, , CLng(0)
    End If

    'Python code
    If pythonCode <> "" Then
        AddDebug "      Region.SaveToFile, Putting Python at " & Seek(f)

        Put #f, , stringToBytes("rPYC")
        Put #f, , CLng(Len(pythonCode))
        Put #f, , stringToBytes(pythonCode)

        Seek #f, Next4bytes(Seek(f))
    End If

    'Color
    AddDebug "      Region.SaveToFile, Putting Color at " & Seek(f)

    Put #f, , stringToBytes("rCOL")
    Put #f, , CLng(4)
    
    Put #f, , CByte(GetRED(color))
    Put #f, , CByte(GetGREEN(color))
    Put #f, , CByte(GetBLUE(color))
    Put #f, , CByte(0)

    '"rNAM" - a descriptive name for the region
    '"rBSE" - whether the region represents a base in a flag game
    '"rNAW" - no antiwarp
    '"rNWP" - no weapons
    '"rNFL" - no flag drops
    '"rAWP" - auto-warp
    '"rPYC" - code to be executed when a player enters or leaves this region
    '"rTIL" - tile data, the definition of the region
    
    'Save unknown sub chunks
    For i = 0 To unknownCount - 1
        
        AddDebug "      Region.SaveToFile, Putting unknown '" & unknownChunk(i).Type & "' at " & Seek(f)
        
        Put #f, , stringToBytes(unknownChunk(i).Type)
        Put #f, , unknownChunk(i).size
        
        If unknownChunk(i).size > 0 Then
            Put #f, , unknownChunk(i).Data
        End If
        
        Seek #f, Next4bytes(Seek(f))
    Next
        
End Function




Private Function MakeTileData(ByRef b() As Byte) As Long
    'Builds the data that will represent the region in the map file
    'Returns the length of the chunk
    
    Dim bytesCount As Long
    
    Dim tilesInRow As Integer
    Dim rowRepeated As Boolean

    Dim i As Integer, k As Integer, j As Integer

    Dim curvalue As Integer

    Dim lastBit As Boolean, newBit As Boolean
    
    Dim rowData() As Integer
    ReDim rowData(256) As Integer
    
    Dim rowDatacount As Long
    
    Dim bmpdata() As Byte
    Dim newByte As Byte
    

    rowDatacount = 0
    
    Dim linelenght As Long
    
    Dim Mask(0 To 7) As Byte
    Mask(0) = 128
    Mask(1) = 64
    Mask(2) = 32
    Mask(3) = 16
    Mask(4) = 8
    Mask(5) = 4
    Mask(6) = 2
    Mask(7) = 1
    
    Call disp.GetBitmapData(bmpdata)
        
    linelenght = disp.GetBitmapLineLenght
    
    For j = 1023 To 0 Step -1
        
'        Call disp.GetBitmapDataLine(bmpdata, j)
        lastBit = (bmpdata(j * linelenght) And Mask(0))
'        lastBit = (bmpdata(j * linelenght) And Mask(0))
'        lastBit = value(0, j)
        curvalue = 0
        
        'k scrolls through each byte of bitmap data
        For k = 0 To 127
            newByte = bmpdata(j * linelenght + k)
            For i = 0 To 7
                'gather row data
    '            newBit = (bmpdata(j * linelenght + (i \ 8)) And Mask(i Mod 8))
                newBit = (newByte And Mask(i))
                
    '            newBit = value(i, j)
                If newBit <> lastBit Then
    
                    If lastBit Then
                        rowData(rowDatacount) = curvalue
                    Else
                        rowData(rowDatacount) = -curvalue
                    End If
                    curvalue = 1
                           
                    rowDatacount = rowDatacount + 1
                    
                    If rowDatacount > UBound(rowData) Then
                        ReDim Preserve rowData(UBound(rowData) + 256)
                    End If
    
                Else
                    curvalue = curvalue + 1
                End If
    
    
                lastBit = newBit
            Next
        Next

        If lastBit Then
            rowData(rowDatacount) = curvalue
        Else
            rowData(rowDatacount) = -curvalue
        End If
        rowDatacount = rowDatacount + 1

        If rowDatacount > UBound(rowData) Then
            ReDim Preserve rowData(UBound(rowData) + 256)
        End If


    Next
        
    
    ReDim b(512) As Byte
    bytesCount = 0
    curvalue = 0

    tilesInRow = 0
    Dim partsInRow As Integer

    i = 0
    While (i < rowDatacount)

        If rowData(i) = -1024 Then
            'empty row
            'check if there are others after
            curvalue = 0
            While (rowData(i) = -1024 And i < rowDatacount)
                curvalue = curvalue + 1
                i = i + 1
'                If i > UBound(rowData) Then GoTo OutOfBounds
            Wend
'OutOfBounds:
            'write the curvalue
            If curvalue <= 32 Then
                '    100nnnnn          - n+1 (1-32) rows of all empty
                
                b(bytesCount) = 127 + curvalue
                bytesCount = bytesCount + 1
                
                If bytesCount > UBound(b) Then
                    ReDim Preserve b(bytesCount + 256) As Byte
                End If
'                parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " empty rows"
            Else
                '    101000nn nnnnnnnn - n+1 (1-1024) rows of all empty

                b(bytesCount) = 160 + (curvalue - 1) \ 256    '128 + 32 ...
                bytesCount = bytesCount + 1
                 
                If bytesCount > UBound(b) Then
                    ReDim Preserve b(bytesCount + 256) As Byte
                End If
                
                b(bytesCount) = (curvalue - 1) Mod 256
                bytesCount = bytesCount + 1
                 
                If bytesCount > UBound(b) Then
                    ReDim Preserve b(bytesCount + 256) As Byte
                End If

'                parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " empty rows"
            End If


        Else
            tilesInRow = 0
            partsInRow = 0
            
            Dim positive As Boolean
                
            Do
                curvalue = Abs(rowData(i))
                tilesInRow = tilesInRow + curvalue
                partsInRow = partsInRow + 1
                
                positive = rowData(i) > 0
                
                'Output the values representing the row
                If curvalue > 32 Then    'more than 32 tiles


                    If positive Then
                        '    011000nn nnnnnnnn - n+1 (1-1024) present tiles in a row
                        b(bytesCount) = 96 + (curvalue - 1) \ 256    '64+32 + (curvalue-1) \ 256

'                        parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " true tiles"
                    Else
                        '    001000nn nnnnnnnn - n+1 (1-1024) empty tiles in a row
                        b(bytesCount) = 32 + (curvalue - 1) \ 256

'                        parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " false tiles"
                    End If
                    
                    bytesCount = bytesCount + 1
                    If bytesCount > UBound(b) Then
                        ReDim Preserve b(bytesCount + 256) As Byte
                    End If
                    
                    b(bytesCount) = (curvalue - 1) Mod 256
                    bytesCount = bytesCount + 1
                    If bytesCount > UBound(b) Then
                        ReDim Preserve b(bytesCount + 256) As Byte
                    End If
                Else

                    If positive Then
                        '    010nnnnn          - n+1 (1-32) present tiles in a row
                        b(bytesCount) = 63 + curvalue

'                        parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " true tiles"
                    Else
                        '    000nnnnn          - n+1 (1-32) empty tiles in a row
                        b(bytesCount) = curvalue - 1

'                        parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " false tiles"
                    End If
                    bytesCount = bytesCount + 1
                    If bytesCount > UBound(b) Then
                        ReDim Preserve b(bytesCount + 256) As Byte
                    End If
                End If

                i = i + 1
            Loop While (tilesInRow < 1024 And i < rowDatacount)





            'Check next rows if there are any repeats
            curvalue = 0

            Do
                rowRepeated = True
                tilesInRow = 0
                j = i
                While (rowRepeated And i < rowDatacount And tilesInRow < 1024)
                    If rowData(i) <> rowData(i - partsInRow) Then
                        rowRepeated = False
                    End If
                    tilesInRow = tilesInRow + Abs(rowData(i - partsInRow))
                    i = i + 1
                Wend
                If rowRepeated Then
                    curvalue = curvalue + 1
                Else
                    i = j
                End If

            Loop While (rowRepeated And i < rowDatacount)

            'save the curvalue
            If curvalue >= 1 And curvalue <= 32 Then
                '    110nnnnn          - repeat last row n+1 (1-32) times

                b(bytesCount) = 191 + curvalue     '128 + 64 + curvalue - 1
                bytesCount = bytesCount + 1
                If bytesCount > UBound(b) Then
                    ReDim Preserve b(bytesCount + 256) As Byte
                End If

'                parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " repeated rows"
            ElseIf curvalue > 32 Then
                '    111000nn nnnnnnnn - repeat last row n+1 (1-1024) times
                
                b(bytesCount) = 224 + (curvalue - 1) \ 256     '128 + 64 + 32 + (curvalue-1) \ 256
                bytesCount = bytesCount + 1
                
                If bytesCount > UBound(b) Then
                    ReDim Preserve b(bytesCount + 256) As Byte
                End If

                b(bytesCount) = (curvalue - 1) Mod 256
                bytesCount = bytesCount + 1
                If bytesCount > UBound(b) Then
                    ReDim Preserve b(bytesCount + 256) As Byte
                End If
                
'                parent.AddDebug "      eLVLdata.SaveRegionTileData, saving " & curvalue & " repeated rows"
            Else
                'no repeats
            End If



        End If




    Wend

    
    If UBound(b) <> bytesCount - 1 Then
        ReDim Preserve b(bytesCount - 1)
    End If
    
    MakeTileData = bytesCount

End Function
