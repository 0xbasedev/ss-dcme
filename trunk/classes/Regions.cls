VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Regions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function CreateFont Lib "GDI32.dll" Alias "CreateFontA" (ByVal h As Long, _
'ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, _
'ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, _
'ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
'Private Declare Function CreateSolidBrush Lib "GDI32.dll" (ByVal crColor As Long) As Long
'Private Declare Function SelectObject Lib "GDI32.dll" ( _
'ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "GDI32.dll" (ByVal hObject As Long) As Long
'Private Declare Function TextOut Lib "GDI32.dll" Alias "TextOutA" (ByVal hdc As Long, _
'ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Private Declare Function Rectangle Lib "GDI32.dll" (ByVal hdc As Long, ByVal x1 As Long, _
'ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'Private Declare Function GetStockObject Lib "GDI32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetROP2 Lib "GDI32.dll" ( _
ByVal hDc As Long, ByVal nDrawMode As Long) As Long
'Private Declare Function SetBkMode Lib "GDI32.dll" ( _
'ByVal hdc As Long, ByVal nBkMode As Long) As Long
'Private Declare Function SetBkColor Lib "GDI32.dll" ( _
'ByVal hdc As Long, ByVal crColor As Long) As Long
'
'Private Const NULL_PEN As Long = &H8
'Private Const OPAQUE As Long = &H2
'Private Const R2_MASKPEN As Long = &H9 ' DPa

Dim Regions() As Region
Dim regionIdx As Integer 'last region index


Dim selregn As Integer 'Index of the selected region



'Display layer for the sum of all regions
Dim maindisp As clsDisplayLayer

Dim parent As frmMain

Dim undoch As Changes

Private Sub Class_Initialize()
    regionIdx = -1
    
    Set maindisp = New clsDisplayLayer
    Call maindisp.Resize(1024, 1024, False)
    
    ReDim Regions(0)
    
'    Set disp = New clsDisplayLayer
    
'    Call disp.Resize(1024, 1024, False)
    
End Sub

Sub setParent(ByRef Main As frmMain)
    Set parent = Main
End Sub



Public Property Get RegionsCount() As Integer
    RegionsCount = regionIdx + 1
End Property





Sub HideAllRegions()
    Dim i As Integer
    
    For i = 0 To regionIdx
    
        Regions(i).visible = False

    Next
    
    Call RedrawAllRegions
        
    Call parent.RedrawRegions(True)
End Sub


Sub SelectRegion(id As Integer, Refresh As Boolean)
    If id >= 0 And id <= regionIdx Then
   
        selregn = id
        
        If Refresh Then
            Call RedrawAllRegions
            Call parent.RedrawRegions(True)
        End If
    End If
End Sub


Sub DrawOnLayer(hDc As Long, screenX As Integer, screenY As Integer, screenWidth As Integer, screenHeight As Integer, firstTileX As Integer, firstTileY As Integer, offsetX As Integer, offsetY As Integer)
    Const MAX_DRAW_WIDTH = 1024
    Const MAX_DRAW_HEIGHT = 768
    
    
    

    
'          Call Regions(selregn).PaintRegionOn(maindisp.hDC, 0, 0, 1023, 1023)
             
    
    
    Dim currenttilew As Integer
    currenttilew = parent.currenttilew
    
    If currenttilew = 1 Then
        BitBlt hDc, screenX, screenY, screenWidth, screenHeight, maindisp.hDc, firstTileX, firstTileY, vbSrcCopy
    Else
        'Add to get the next complete tile
        screenWidth = screenWidth + offsetX
        screenHeight = screenHeight + offsetY
        If screenWidth Mod currenttilew <> 0 Then screenWidth = screenWidth + currenttilew - (screenWidth Mod currenttilew)
        If screenHeight Mod currenttilew <> 0 Then screenHeight = screenHeight + currenttilew - (screenHeight Mod currenttilew)
        
        'Sometimes, if the pic is too large, the buffer won't be able to handle it, and the stretchblt will fail
        'Apparently stretchblt's buffer for a scanline is maxed at 64k
        'So let's do 500x500 sections
        Dim i As Integer, j As Integer
        Dim drawWidth As Long, drawHeight As Long
        Dim curtilex As Long, curtiley As Long
        

                
'130           StretchBlt hDC, screenX, screenY, screenWidth, screenHeight, maindisp.hDC, curtilex, curtiley, screenWidth / currenttilew, screenHeight / currenttilew, vbSrcCopy



        curtiley = firstTileY
        For j = 0 To screenHeight Step MAX_DRAW_HEIGHT

            curtilex = firstTileX

            For i = 0 To screenWidth Step MAX_DRAW_WIDTH

                If screenWidth - i < MAX_DRAW_WIDTH Then drawWidth = screenWidth - i Else drawWidth = MAX_DRAW_WIDTH
                If screenHeight - j < MAX_DRAW_HEIGHT Then drawHeight = screenHeight - j Else drawHeight = MAX_DRAW_HEIGHT

                StretchBlt hDc, screenX + i - offsetX, screenY + j - offsetY, drawWidth, drawHeight, maindisp.hDc, curtilex, curtiley, drawWidth \ currenttilew, drawHeight \ currenttilew, vbSrcCopy

'                    frmGeneral.Label6.Caption = screenX + i - offsetX & "," & screenY + j - offsetY & "," & drawWidth & "," & drawHeight & "," & drawWidth \ currenttilew & "," & drawHeight \ currenttilew
              
                curtilex = curtilex + (MAX_DRAW_WIDTH \ currenttilew)
            Next

            curtiley = curtiley + (MAX_DRAW_HEIGHT \ currenttilew)
        Next
'        frmGeneral.Label6.Caption = "Stretch: " & StretchBlt(hdc, screenX, screenY, screenWidth, screenHeight, maindisp.hdc, firstTileX, firstTileY, screenWidth / currenttilew, screenHeight / currenttilew, vbSrcCopy)
    End If
End Sub


Private Sub PaintRegionArea(regionID As Integer, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, selected As Boolean)
          
    maindisp.TextColor = vbBlack
    SetBkColor maindisp.hDc, Regions(regionID).color
   
    If selected Then
        TransparentBlt maindisp.hDc, x1, y1, x2 - x1 + 1, y2 - y1 + 1, Regions(regionID).disphDC, x1, y1, vbBlack
    Else
'        TransAlphaBlt maindisp.hDC, x1, y1, x2 - x1 + 1, y2 - y1 + 1, Regions(regionID).disp.hDC, x1, y1, vbBlack, 64
        BitBlt maindisp.hDc, x1, y1, x2 - x1 + 1, y2 - y1 + 1, Regions(regionID).disphDC, x1, y1, vbSrcPaint
    End If

End Sub

Private Sub RedrawAllRegionsArea(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    Dim i As Integer
    
    Call maindisp.EraseArea(x1, y1, x2 + 1, y2 + 1)
    
    For i = 0 To regionIdx
      If i <> selregn Then
          If Regions(i).visible Then
              Call PaintRegionArea(i, x1, y1, x2, y2, False)
          End If
      End If
    Next
    
    If selregn >= 0 And selregn <= regionIdx Then
      If Regions(selregn).visible Then
          Call PaintRegionArea(selregn, x1, y1, x2, y2, True)
      End If
    
'            Call Regions(selregn).disp.SaveToFile(Directory_Temp & "\temp_regn.bmp", False)
'            Call maindisp.SaveToFile(Directory_Temp & "\temp_main.bmp", False)
    End If
'20        For i = 0 To regionIdx
'30            If Regions(i).visible Then Call PaintRegionArea(i, x1, y1, x2, y2)
'40        Next
    'for each region
        'bitblt on maindisp with mergepaint
    'next
    
'    Dim hdc As Long, hOldBrush As Long, hBrush As Long, hOldPen As Long
'    Dim tick As Long
'    hdc = disp.hdc
'
'    If add Then
'        SetROP2 hdc, R2_MERGEPEN
'
'        hOldPen = SelectObject(hdc, GetStockObject(NULL_PEN))
'
'        hBrush = CreateSolidBrush(color)
'        hOldBrush = SelectObject(hdc, hBrush)
'
'
'        Call Rectangle(hdc, x1, y1, x2 + 2, y2 + 2)
'
'    '                    Call DrawFilledRectangle(hDC, screenX, screenY, screenX + currenttilew, screenY + currenttilew, .color)
'        Call SelectObject(hdc, hOldBrush)
'        Call DeleteObject(hBrush)
'        Call SelectObject(hdc, hOldPen)
'
'    Else
'        'If there is only one region, we can simply draw black on the region
'
'        If regionIdx = 0 Then
'
'            SetROP2 hdc, R2_BLACK
'
'            hOldPen = SelectObject(hdc, GetStockObject(NULL_PEN))
'
'            hBrush = CreateSolidBrush(color)
'
'            hOldBrush = SelectObject(hdc, hBrush)
'
'
'            Call Rectangle(hdc, x1, y1, x2 + 2, y2 + 2)
'
'        '                    Call DrawFilledRectangle(hDC, screenX, screenY, screenX + currenttilew, screenY + currenttilew, .color)
'            Call SelectObject(hdc, hOldBrush)
'            Call DeleteObject(hBrush)
'            Call SelectObject(hdc, hOldPen)
'        ElseIf CLng(x2 - x1) * CLng((y2 - y1)) > 5000 Then
'            'If the area is large enough, use the clsPic
'
'
'            tick = GetTickCount
'            Call RedrawAllRegionsArea(x1, y1, x2, y2)
'
'            frmGeneral.Label6.Caption = ">; " & CLng(x2 - x1) * CLng((y2 - y1)) & "tiles ; " & GetTickCount - tick & "ms ; " & CDbl(GetTickCount - tick) / (CLng(x2 - x1) * CLng((y2 - y1))) & "ms/t"
'        Else
'            'We have to calculate each pixel individually :(
'
'            Dim i As Long, j As Long, r As Integer
'            Dim finalColor As Long
'
'
'            SetROP2 hdc, R2_COPYPEN
'
'
'            tick = GetTickCount
'
'            For j = y1 To y2
'                For i = x1 To x2
'                    finalColor = 0
'                    For r = 0 To regionIdx
'                        If Regions(r).bitfield.valueL(i, j) Then If Regions(r).visible Then finalColor = finalColor Or Regions(r).color
'                    Next
'
'                    Call SetPixel(hdc, i, j, finalColor)
'                Next
'            Next
'
'            frmGeneral.Label6.Caption = "<; " & CLng(x2 - x1) * CLng((y2 - y1)) & "tiles ; " & GetTickCount - tick & "ms ; " & CDbl(GetTickCount - tick) / (CLng(x2 - x1) * CLng((y2 - y1))) & "ms/t"
'        End If
'    End If
End Sub


Sub RedrawAllRegions()
    Call RedrawAllRegionsArea(0, 0, 1023, 1023)
End Sub

Sub AddToRegion(regionID As Integer, X As Integer, Y As Integer, ByRef undoch As Changes)
    If regionID >= 0 And regionID <= regionIdx Then
       
        Call Regions(regionID).AddTile(X, Y, undoch, regionID)

    End If
End Sub

Sub RemoveFromRegion(regionID As Integer, X As Integer, Y As Integer, undoch As Changes)
    If regionID >= 0 And regionID <= regionIdx Then
        
        Call Regions(regionID).RemoveTile(X, Y, undoch, regionID)
        
    End If
End Sub

Sub AddRectangleToRegion(regionID As Integer, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByRef undoch As Changes)
    Dim tick As Long
    tick = GetTickCount
    
    Dim i As Integer, j As Integer

    If regionID >= 0 And regionID <= regionIdx Then
        If x1 > x2 Then Swap x1, x2
        If y1 > y2 Then Swap y1, y2
        
        Call Regions(regionID).AddRectangle(x1, y1, x2, y2, undoch, regionID)
        
'60            Call RedrawAllRegionsArea(x1, y1, x2, y2)
        If Regions(regionID).visible Then Call PaintRegionArea(regionID, x1, y1, x2, y2, True)

    End If
    
'80        frmGeneral.Label7.Caption = "Add: " & GetTickCount - tick
End Sub

Sub RemoveRectangleFromRegion(regionID As Integer, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByRef undoch As Changes)
    Dim i As Integer, j As Integer
    Dim tick As Long
    tick = GetTickCount
    
    If regionID >= 0 And regionID <= regionIdx Then
        If x1 > x2 Then Swap x1, x2
        If y1 > y2 Then Swap y1, y2
        
        Call Regions(regionID).RemoveRectangle(x1, y1, x2, y2, undoch, regionID)
        
        If Regions(regionID).visible Then Call RedrawAllRegionsArea(x1, y1, x2, y2)

    End If
    
'80        frmGeneral.Label7.Caption = "Remove: " & GetTickCount - tick
End Sub

Friend Sub setRegion(ByRef r As Region, Index As Integer)
    If regionIdx < 0 Then
        ReDim Regions(8)
    ElseIf Index >= UBound(Regions) Then
        ReDim Preserve Regions(UBound(Regions) + 8)
    End If
    
    If Regions(Index) Is Nothing Then
        regionIdx = regionIdx + 1
    End If
    Set Regions(Index) = r
End Sub

'sets all regions
'Friend Sub setRegions(ByRef r() As Region, rIdx As Integer)
'    Dim i As Integer
'
'    If rIdx >= UBound(Regions) Then
'        ReDim Preserve Regions(rIdx + 8)
'    End If
'
'    For i = 0 To rIdx
'        Set Regions(i) = r(i)
'    Next
'    For i = rIdx + 1 To regionIdx
'        Set Regions(i) = Nothing
'    Next
'    regionIdx = rIdx
'End Sub

Friend Function getRegion(Index As Integer) As Region
    Set getRegion = Regions(Index)
End Function

'Friend Function getRegions() As Region()
'    getRegions = Regions
'End Function

Function getRegionIndex() As Integer
    getRegionIndex = regionIdx
End Function



'Sub BuildRegionTiles()
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'
'    parent.picRegionColors.height = (regionIdx + 1) * TILEW * 2
'
'    For k = 0 To regionIdx
'        For j = 0 To TILEW * 2
'            For i = 0 To TILEW * 2
'                Call SetPixel(parent.picRegionColors.hDC, i, k * TILEW * 2 + j, Regions(k).color)
'            Next
'        Next
'    Next
'End Sub


Function getRegionsNamesAt(X As Integer, Y As Integer) As String
    Dim i As Integer, count As Boolean
    
    For i = 0 To regionIdx
        If Regions(i).value(X, Y) Then
                
            
            If count Then
                getRegionsNamesAt = getRegionsNamesAt & ", " & Regions(i).name
            Else
                getRegionsNamesAt = Regions(i).name
                count = True
            End If
            
        End If
    Next
End Function

Function isRegionAt(X As Integer, Y As Integer) As Integer
    Dim i As Integer
    For i = regionIdx To 0 Step -1
        If Regions(i).value(X, Y) Then
            isRegionAt = i
            Exit Function
        End If
    Next

    isRegionAt = -1
End Function

Function getRegionName(id As Integer) As String
    getRegionName = Regions(id).name
End Function

Sub setRegionName(id As Integer, newname As String)
    Regions(id).name = newname
End Sub
Function getRegionColor(id As Integer) As Long
    getRegionColor = Regions(id).color
End Function

Function getRegionIsVisible(id As Integer) As Boolean
    getRegionIsVisible = Regions(id).visible
End Function

Function getRegionPython(id As Integer) As String
    getRegionPython = Regions(id).pythonCode
End Function

Sub SetRegionPython(id As Integer, str As String)
    Regions(id).pythonCode = str
End Sub

'checks if at least one of the regions is visible
Function HaveVisibleRegions() As Boolean
    HaveVisibleRegions = False
    Dim i As Integer
    For i = 0 To regionIdx
        If Regions(i).visible Then
            HaveVisibleRegions = True
            Exit Function
        End If
    Next
End Function

Function CountVisibleRegions() As Integer
    CountVisibleRegions = 0
    Dim i As Integer
    For i = 0 To regionIdx
        If Regions(i).visible Then CountVisibleRegions = CountVisibleRegions + 1
    Next
End Function




'Sub DrawRegions(Optional step As Integer = 1)
'    If regionIdx = -1 Then Exit Sub
'
'    'CALCULATE BOUNDARIES OF PICLEVEL
'    Dim lbx As Integer    'lowestval for x
'    Dim hbx As Integer    'highestval for x
'    Dim lby As Integer    'lowestval for y
'    Dim hby As Integer    'highestval for y
'
'    lbx = parent.Hscr.value \ parent.currenttilew
'    hbx = (parent.Hscr.value + parent.picPreview.width) \ parent.currenttilew
'    lby = parent.Vscr.value \ parent.currenttilew
'    hby = (parent.Vscr.value + parent.picPreview.height) \ parent.currenttilew
'
'    If lbx < 0 Then lbx = 0
'    If hbx > 1023 Then hbx = 1023
'    If lby < 0 Then lby = 0
'    If hby > 1023 Then hby = 1023
'
'    Dim i As Long, j As Long
'
'    For j = lby To hby Step step
'        For i = lbx To hbx Step step
''            Call DrawRegionOn(i, j, lbx, lby)
'        Next
'    Next
'End Sub

'Function DrawRegionOn(hdc As Long, X As Integer, Y As Integer, screenX As Integer, screenY As Integer) As Boolean
'    Dim i As Integer, currenttilew As Integer
'    currenttilew = parent.currenttilew
'
'    SetROP2 hdc, R2_LAST
'
'    Dim hOldBrush As Long, hBrush As Long, hOldPen As Long
'
'    hOldPen = SelectObject(hdc, GetStockObject(NULL_PEN))
'
'
'    For i = regionIdx To 0 Step -1
'        With Regions(i)
'            If .visible Then
'                If .bitfield.value(X, Y) Then
'                    hBrush = CreateSolidBrush(.color)
'                    hOldBrush = SelectObject(hdc, hBrush)
'
'                    DrawRegionOn = True
'                    Call Rectangle(hdc, screenX, screenY, screenX + currenttilew + 1, screenY + currenttilew + 1)
'
''                    Call DrawFilledRectangle(hDC, screenX, screenY, screenX + currenttilew, screenY + currenttilew, .color)
'                    Call SelectObject(hdc, hOldBrush)
'                    Call DeleteObject(hBrush)
'                End If
'            End If
'        End With
'    Next
'
'    Call SelectObject(hdc, hOldPen)
'
'
'End Function


'Private Function getDefaultRegion() As MAPregionType
'    With getDefaultRegion
'        .isAutoWarp = False
'        .autowarpX = 0
'        .autowarpY = 0
'        .autowarpArena = ""
'        .isBase = False
'        .isNoAntiwarp = False
'        .isNoFlagDrop = False
'        .isNoWeapon = False
'        .name = "New Region"
'        .pythonCode = ""
'        .visible = True
'    End With
'End Function

'creates new empty region and returns the index of the region
Function NewRegion(Optional name As String = "") As Integer
    If regionIdx = -1 Then
        ReDim Regions(0)
    ElseIf regionIdx + 1 >= UBound(Regions) Then
        ReDim Preserve Regions(UBound(Regions) + 8)
    End If
    
    
    regionIdx = regionIdx + 1
    
    Set Regions(regionIdx) = New Region
    
'    Call Regions(regionIdx).bitfield.InitArray(1024, 1024, 0)
    'ReDim Regions(regionIdx).bitfield(1023, 1023)

    If name = "" Then
        Dim i As Integer
        i = 1
        While regionNameExists("New Region " & CStr(i))
            i = i + 1
        Wend
        name = "New Region " & CStr(i)
    End If


    Randomize
    
    Dim cR As Integer, cG As Integer, cB As Integer

    Select Case regionIdx Mod 6
        Case 0
            cR = 130 + Rnd() * 125
            cG = Rnd() * 100
            cB = Rnd() * 100
        Case 1
            cR = Rnd() * 100
            cG = 130 + Rnd() * 125
            cB = Rnd() * 100
        Case 2
            cR = Rnd() * 100
            cG = Rnd() * 100
            cB = 130 + Rnd() * 125
        Case 3
            cR = 130 + Rnd() * 125
            cG = 130 + Rnd() * 125
            cB = Rnd() * 100
        Case 4
            cR = Rnd() * 100
            cG = 130 + Rnd() * 125
            cB = 130 + Rnd() * 125
        Case 5
            cR = 130 + Rnd() * 125
            cG = Rnd() * 100
            cB = 130 + Rnd() * 125
    End Select
    
    Call Regions(regionIdx).InitRegion(name, RGB(cR, cG, cB))

'    Call BuildRegionTiles

    NewRegion = regionIdx
End Function

Sub DeleteRegion(Index As Integer)
    Dim i As Integer
    For i = Index + 1 To regionIdx
        Set Regions(i - 1) = Regions(i)
        
    Next
        
        
    Set Regions(regionIdx) = Nothing

    regionIdx = regionIdx - 1

    Call RedrawAllRegions
    
'    Call BuildRegionTiles

    Call parent.RedrawRegions(True)
End Sub

Sub ToggleVisible(Index As Integer, Refresh As Boolean)
    Regions(Index).visible = Not Regions(Index).visible
    
    Call RedrawAllRegions
    
    If Refresh Then
        
        Call parent.RedrawRegions(True)
'        Call parent.UpdateLevel
    End If
End Sub

Sub SetColor(Index As Integer, color As Long)
    Regions(Index).color = color
    
'20        Call RedrawAllRegions
'
'30        If Regions(Index).visible Then
'40            Call parent.RedrawRegions(True)
'50        End If
End Sub

Function regionNameExists(str As String) As Boolean
    If regionIdx = -1 Then
        regionNameExists = False
    Else
        Dim i As Integer
        For i = 0 To regionIdx
            If Regions(i).name = str Then
                regionNameExists = True
                Exit Function
            End If
        Next
    End If
End Function


Private Sub Swap(ByRef val1 As Integer, ByRef val2 As Integer)
    Dim tmp As Integer
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub




Private Sub Class_Terminate()
    Set parent = Nothing
'    Call maindisp.Resize(1, 1, False)
    Set maindisp = Nothing
    
    
    
    Dim i As Integer
    For i = 0 To UBound(Regions)
        Set Regions(i) = Nothing
    Next
End Sub
