VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LVZData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum CursorState
    S_Nothing = 0
    S_SelectingSingle
    S_SelectingArea
    S_Moving
End Enum

Dim lvz() As LVZstruct

Dim lvzcount As Long

Dim defaultLayer As LVZLayerEnum
Dim defaultDisplay As LVZModeEnum

Dim totalImageDefinitions As Integer

Dim parent As frmMain
Dim undoch As Changes

Dim State As CursorState

Dim firstX As Integer, firstY As Integer 'First coordinates we clicked (in pixels, relative to 0,0)
Dim lastX As Integer, lastY As Integer


Dim c_Snap As Integer 'Snap for moving objects in pixels



Dim selectedObjects() As MapObjectRef
Dim selectedObjectsCount As Long

'Properties of the selected set of objects
'These will be set to -1 if multiple values are found in the selection
Dim selLayer As LVZLayerEnum
Dim selDisplayMode As LVZModeEnum
Dim selObjectID As Integer
Dim selDisplayTime As Long

Dim selCoordX As Integer 'Set to -32768 if not all the same
Dim selCoordY As Integer
   
Private Const MULTIPLE_VALUES = -32768
     
Public dontCallSelectionChange As Boolean 'used when changing combo boxes by code

Dim c_maxidx As Integer

'Dim picLVZImagesZoomPxl() As clsDisplayLayer
'Dim picLVZImagesZoom1_4() As clsDisplayLayer
'Dim picLVZImagesZoom1_2() As clsDisplayLayer
Dim picLVZImages() As clsDisplayLayer
Dim picLVZImagesZoom() As clsDisplayLayer
'Dim picLVZImagesZoom2_1() As clsDisplayLayer

Dim picUsed() As Boolean

Dim picLVZImagesLibrary() As clsDisplayLayer



Public Property Get pichDCzoom(idx As Integer, zoom As Single) As Long
    If picUsed(idx) Then
      If zoom = 1 Then
          pichDCzoom = picLVZImages(idx).hDC
      Else
          pichDCzoom = picLVZImagesZoom(idx).hDC
      End If
'20            If zoom = 1 Then
'30                pichDCzoom = picLVZImages(idx).hDC
'40            ElseIf zoom = 1 / 2 Then
'50                pichDCzoom = picLVZImagesZoom1_2(idx).hDC
'60            ElseIf zoom = 1 / 4 Then
'70                pichDCzoom = picLVZImagesZoom1_4(idx).hDC
'80            ElseIf zoom = 2 Then
'90                pichDCzoom = picLVZImagesZoom2_1(idx).hDC
'100           Else
'110               pichDCzoom = picLVZImagesZoomPxl(idx).hDC
'120           End If

    Else
        Err.Raise ERR_UNINITIALIZED_LVZ_DC, Me, "Tried to use a non-existant device context for LVZ image " & idx
    End If
End Property

Public Property Get pichDC(idx As Integer) As Long
    If picUsed(idx) Then
        pichDC = picLVZImages(idx).hDC
    Else
        Err.Raise ERR_UNINITIALIZED_LVZ_DC, Me, "Tried to use a non-existant device context for LVZ image " & idx
    End If
End Property

Public Property Get pichDClib(idx As Integer) As Long
    If picUsed(idx) Then
        pichDClib = picLVZImagesLibrary(idx).hDC
    Else
        Err.Raise ERR_UNINITIALIZED_LVZ_DC, Me, "Tried to use a non-existant device context for LVZ image thumbnail " & idx
    End If
End Property

'Dim visibleObjectsOver() As MapObjectRef
'Dim visibleObjectsOverCount As Long
'Dim visibleObjectsUnder() As MapObjectRef
'Dim visibleObjectsUnderCount As Long
'
'Dim visibleArea As RECT

Private Sub ResizePicLVZArrays(maxidx As Integer, b_preserve As Boolean)
    If b_preserve Then
'        ReDim Preserve picLVZImagesZoomPxl(maxidx)
'        ReDim Preserve picLVZImagesZoom1_4(maxidx)
'        ReDim Preserve picLVZImagesZoom1_2(maxidx)
        ReDim Preserve picLVZImages(maxidx)
        ReDim Preserve picLVZImagesZoom(maxidx)
'        ReDim Preserve picLVZImagesZoom2_1(maxidx)
        
        ReDim Preserve picUsed(maxidx)
        
        ReDim Preserve picLVZImagesLibrary(maxidx)
    Else
        Dim i As Integer
        For i = 0 To c_maxidx
'            Set picLVZImagesZoomPxl(i) = Nothing
'            Set picLVZImagesZoom1_4(i) = Nothing
'            Set picLVZImagesZoom1_2(i) = Nothing
            Set picLVZImages(i) = Nothing
            Set picLVZImagesZoom(i) = Nothing
'            Set picLVZImagesZoom2_1(i) = Nothing
            Set picLVZImagesLibrary(i) = Nothing
        Next
        
'        ReDim picLVZImagesZoomPxl(maxidx)
'        ReDim picLVZImagesZoom1_4(maxidx)
'        ReDim picLVZImagesZoom1_2(maxidx)
        ReDim picLVZImages(maxidx)
        ReDim picLVZImagesZoom(maxidx)
'        ReDim picLVZImagesZoom2_1(maxidx)
        
        ReDim picUsed(maxidx)
        
        ReDim picLVZImagesLibrary(maxidx)
    End If
        
    c_maxidx = maxidx
End Sub



Private Sub Class_Initialize()
'Initializes all dynamic arrays of lvz and images of lvz and map objects of lvz etc etc
'Sets default layer and display mode
'@Post all arrays within lvz are ready to use
'@Post selected layer = 3 , selected display mode = 0

    ReDim lvz(0)
    lvzcount = 0
    ReDim lvz(0).files(0)
    lvz(0).filecount = 0
    ReDim lvz(0).mapobjects(0)
    lvz(0).mapObjectCount = 0
    

    
    ReDim lvz(0).screenobjects(0)
    lvz(0).ScreenobjectCount = 0
    ReDim lvz(0).imagedefinitions(0)
    lvz(0).ImageDefinitionCount = 0
    totalImageDefinitions = 0
    
    defaultLayer = lyr_AfterTiles
    defaultDisplay = md_ShowAlways
    
'    SelLVZIdx = -1
'    selObjIdx = -1
    
    c_Snap = 1
    
'    ReDim visibleObjectsOver(0)
'    visibleObjectsOverCount = 0
'    ReDim visibleUnderObjects(0)
'    visibleObjectsUnderCount = 0
    
    ReDim selectedObjects(0)
    selectedObjectsCount = 0
    
    c_maxidx = -1
    Call ResizePicLVZArrays(0, False)
'    Call DoSelect
End Sub



Private Sub RedrawLVZLayers()
    Call parent.RedrawLvzUnderLayer(False)
    Call parent.RedrawLvzOverLayer(True)
End Sub




Sub AddToSelection(lvzidx As Integer, objidx As Long)
    
    With lvz(lvzidx).mapobjects(objidx)
        If Not .selected Then
            selectedObjects(selectedObjectsCount).lvzidx = lvzidx
            selectedObjects(selectedObjectsCount).objidx = objidx
            
            
            
            selectedObjectsCount = selectedObjectsCount + 1
            
            If UBound(selectedObjects) < selectedObjectsCount Then
                ReDim Preserve selectedObjects(selectedObjectsCount + 16)
            End If
    
            .selected = True
            
            lvz(lvzidx).totalSelected = lvz(lvzidx).totalSelected + 1
            
            Call UpdateSelectedProperties(selectedObjectsCount - 1)

        End If
    End With
End Sub

Private Sub UpdateSelectedProperties(firstidx As Long)
    Dim i As Long
    
    For i = firstidx To selectedObjectsCount - 1
        With lvz(selectedObjects(i).lvzidx).mapobjects(selectedObjects(i).objidx)
            If firstidx = 0 Then
               'First object added
                selLayer = .layer
                selDisplayMode = .mode
                selDisplayTime = .displayTime
                selObjectID = .objectID
                selCoordX = .X
                selCoordY = .Y
            Else
                If .layer <> selLayer Then selLayer = -1
                If .mode <> selDisplayMode Then selDisplayMode = -1
                If .displayTime <> selDisplayTime Then selDisplayTime = -1
                If .objectID <> selObjectID Then selObjectID = -1
                If .X <> selCoordX Then selCoordX = MULTIPLE_VALUES
                If .Y <> selCoordY Then selCoordY = MULTIPLE_VALUES
                
            End If
        End With
    Next
End Sub


Sub RemoveFromSelection(lvzidx As Integer, objidx As Long)
    Dim i As Long
    
    With lvz(lvzidx)
        If Not .mapobjects(objidx).selected Then Exit Sub
    
        .mapobjects(objidx).selected = False
        .totalSelected = .totalSelected - 1
        
    End With
    
    selectedObjectsCount = selectedObjectsCount - 1
    
    'Search for this map object in our list
    For i = 0 To selectedObjectsCount
        If selectedObjects(i).lvzidx = lvzidx Then
            If selectedObjects(i).objidx = objidx Then
                'We found it
                'Now, take everything after it and memcopy it back
                If i < selectedObjectsCount Then
                    CopyMemory ByVal VarPtr(selectedObjects(i)), ByVal VarPtr(selectedObjects(i + 1)), (selectedObjectsCount - i) * LenB(selectedObjects(i))
                End If
                
                Exit For
            End If
        End If

    Next

  Call UpdateSelectedProperties(0)
End Sub


Sub EmptySelection()
    Dim i As Integer
    For i = 0 To selectedObjectsCount
        lvz(selectedObjects(i).lvzidx).mapobjects(selectedObjects(i).objidx).selected = False
    Next
    selectedObjectsCount = 0
    
    For i = 0 To lvzcount - 1
        lvz(i).totalSelected = 0
    Next
End Sub



Sub setParent(Main As frmMain)
    Set parent = Main
End Sub

Friend Function getLVZData() As LVZstruct()
'Retrieves the entire lvz struct
'@Returns lvz data
    getLVZData = lvz
End Function

Function getLVZname(lvzIndex As Integer) As String
'Retrieve the name of the lvz
'@Returns lvz name
    getLVZname = lvz(lvzIndex).name
End Function

Function getLVZpath(lvzIndex As Integer) As String
'Retrieve the path of the lvz
'@Returns lvz path
    getLVZpath = lvz(lvzIndex).path
End Function

Function getLVZCount() As Long
'Retrieves the number of lvz's
'@Returns numer of lvz's
    getLVZCount = lvzcount
End Function

Function getFileCount(lvzIndex As Integer) As Long
'Retrieves the number of files in the given lvz
'@Pre lvzIndex is within valid boundaries
'@Params lvzIndex: index of lvz
'@Returns number of files
    getFileCount = lvz(lvzIndex).filecount
End Function

Function getImageDefinitionCount(lvzIndex As Integer) As Integer
'Retrieves the number of images in the given lvz
'@Pre lvzIndex is within valid boundaries
'@Params lvzIndex: index of lvz
'@Returns numer of images

    getImageDefinitionCount = lvz(lvzIndex).ImageDefinitionCount
End Function

Function getMapobjectCount(lvzIndex As Integer) As Long
'Retrieves the number of map objects in a given lvz
'@Pre lvzIndex is within valid boundaries
'@Params lvzIndex: index of lvz
'@Returns map object count
    getMapobjectCount = lvz(lvzIndex).mapObjectCount
End Function

Function getScreenobjectCount(lvzIndex As Integer) As Long
'Retrieves the number of screen objects in a given lvz
'@Pre lvzIndex is within valid boundaries
'@Params lvzIndex: index of lvz
'@Returns screen object count
    getScreenobjectCount = lvz(lvzIndex).ScreenobjectCount
End Function

Friend Function getFileData(lvzIndex As Integer, fileIndex As Long) As LVZFileStruct
'Retrieves a file entry of a specified lvz
'@Pre lvzIndex and fileIndex is within valid boundaries
'@Params
'   lvzIndex: index of lvz
'   fileindex: index of lvz
'@Returns file data of an lvz

    getFileData = lvz(lvzIndex).files(fileIndex)
End Function

Friend Function getMapObject(lvzIndex As Integer, mapobjIndex As Long) As LVZMapObject
'Retrieves a map object entry of a specified lvz
    getMapObject = lvz(lvzIndex).mapobjects(mapobjIndex)
End Function

Friend Function getScreenObject(lvzIndex As Integer, scrobjIndex As Long) As LVZScreenObject
'Retrieves a screen object entry of a specified lvz
    getScreenObject = lvz(lvzIndex).screenobjects(scrobjIndex)
End Function

Friend Function getImageDefinition(lvzIndex As Integer, imgdefIndex As Integer) As LVZImageDefinition
'Retrieves an image entry of a specified lvz
    getImageDefinition = lvz(lvzIndex).imagedefinitions(imgdefIndex)
End Function

Sub UpdateImageAnimation(lvzidx As Integer, imgidx As Integer, tick As Long, frame As Long)
    lvz(lvzidx).imagedefinitions(imgidx).CurrentFrame = frame
    lvz(lvzidx).imagedefinitions(imgidx).lastFrameChange = tick
End Sub

Friend Function getLVZ(lvzidx As Integer) As LVZstruct
'Retrieves an lvz struct
    getLVZ = lvz(lvzidx)
End Function

Friend Sub setLVZ(editedlvz As LVZstruct, lvzidx As Integer)
'Sets an lvz struct to the given index
    lvz(lvzidx) = editedlvz
End Sub

Friend Sub setLVZData(tmplvz() As LVZstruct, count As Integer)
'Sets an entire array of lvz's defined
    lvz = tmplvz
    lvzcount = count
End Sub

Function moveLVZFile(flIdx As Long, frmIdx As Integer, toIdx As Integer) As Boolean
'Move an lvz file entry from one lvz to another

    On Error GoTo errh
    
    Dim fl As LVZFileStruct
    'fl = tmplvz(frmIdx).files(flIdx)
    'tmplvz =... ?
    
    
    If AddFileToLVZ(toIdx, fl.path) Then
        'add file succeeded, remove old one
        Call removeFileFromLVZ(frmIdx, flIdx)
    Else
        moveLVZFile = False
        Exit Function
    End If
    
    moveLVZFile = True
    
Exit Function
errh:
    moveLVZFile = False
End Function

Function AddLVZ(name As String) As Boolean
'Adds an empty lvz with a given name

    If lvzcount > UBound(lvz) Then
        ReDim Preserve lvz(UBound(lvz) + 5)
    End If
    
    If getIndexOfLVZ(name) <> -1 Then
        'another lvz with the same name exists
        AddLVZ = False
        Exit Function
    End If
    
    With lvz(lvzcount)
    
        ReDim .files(0)
        .filecount = 0
        .name = name
        
        .path = GetPathTo(parent.activeFile) & name
        
        ReDim .mapobjects(0)
        .mapObjectCount = 0
        ReDim .screenobjects(0)
        .ScreenobjectCount = 0
        ReDim .imagedefinitions(0)
        .ImageDefinitionCount = 0
        
        Dim i As Integer
        For i = lyr_BelowAll To lyr_TopMost
            .nextIndexFor(i) = 0
        Next
    
    End With
    
    lvzcount = lvzcount + 1
    
    
    AddLVZ = True
End Function

Function AddFileToLVZ(lvzidx As Integer, path As String) As Long
'Adds a file with given path to the given lvz
    With lvz(lvzidx)
    
        If .filecount > UBound(.files) Then
            ReDim Preserve .files(UBound(.files) + 5)
        End If
        
        If getIndexOfFile(lvzidx, path) <> -1 Then
            'another file in the lvz with the same name exists
            AddFileToLVZ = -1
            Exit Function
        End If
        
        .files(.filecount).path = path
        .files(.filecount).Type = getLVZFileType(path)
        
        AddFileToLVZ = .filecount
        
        .filecount = .filecount + 1
    
    End With
    
End Function


Private Function GetIndexForMapobject(lvzidx As Integer, ByRef obj As LVZMapObject) As Long
    'Finds where to place an object, depending on its layer, using successive approximations
    GetIndexForMapobject = lvz(lvzidx).nextIndexFor(obj.layer)

End Function

Private Sub InsertMapObjectAtIdx(lvzidx As Integer, ByRef obj As LVZMapObject, objidx As Long)
'    Dim objidxMax As Long
    
    With lvz(lvzidx)
'        objidxMax = .mapObjectCount
        
        'Insert it
        ' First, copymemory from objidx to the last
        If objidx < .mapObjectCount And .mapObjectCount > 0 Then
            CopyMemory ByVal VarPtr(.mapobjects(objidx + 1)), _
                        ByVal VarPtr(.mapobjects(objidx)), _
                        LenB(obj) * (.mapObjectCount - objidx)
        End If
        ' Then, insert at objidx
        .mapobjects(objidx) = obj
        
'        .lastOfLayer(obj.layer) = objIdx
        
        
        'Increment the target indexes
        Dim i As Integer
        For i = obj.layer To lyr_TopMost
            .nextIndexFor(i) = .nextIndexFor(i) + 1
        Next
    End With
    
End Sub






Friend Function AddMapObjectToLVZ(lvzidx As Integer, obj As LVZMapObject) As Long
'Adds a given map object to the given lvz
'Returns the index of the added item

    With lvz(lvzidx)
    
        If .mapObjectCount > UBound(.mapobjects) Then
            ReDim Preserve .mapobjects(.mapObjectCount + 256)
        End If
        
        'Find the index where to insert the object
        'Sorted by: Layer
        'The object should be placed as the last of the same layer.
        'ex: if objects 10 to 20 are of the same layer as our new object, object should be placed as 21
        '    objects 21 to <last> will be copied to 22 to <last>
        
        Dim objidx As Long
        
        objidx = GetIndexForMapobject(lvzidx, obj)
        
        Call InsertMapObjectAtIdx(lvzidx, obj, objidx)
        
    '    frmGeneral.Label7.Caption = "Index '" & objIdx & "/" & objidxMax & "' found in " & n & " iterations (" & GetTickCount - tick & " ms)"
        
        
        AddMapObjectToLVZ = objidx
            
        .mapObjectCount = .mapObjectCount + 1
        
        'Adjust selection indexes
        
        If .totalSelected > 0 Then
            Dim i As Long
            For i = 0 To selectedObjectsCount - 1
                With selectedObjects(i)
                    If .lvzidx = lvzidx Then
                        If .objidx >= objidx Then .objidx = .objidx + 1
                    End If
                End With
            Next
        End If

    End With
    
End Function



Friend Function AddScreenObjectToLVZ(lvzidx As Integer, obj As LVZScreenObject) As Long
'Adds a given screen object to the given lvz
'@Pre: index is within valid boundaries
'@Params
'    index: index of lvz
'    obj: data of screen object
'@Post: screen object is added, screenobject count is increased by 1
    
    With lvz(lvzidx)
        If .ScreenobjectCount > UBound(.screenobjects) Then
            ReDim Preserve .screenobjects(.ScreenobjectCount + 64)
        End If
            
        .screenobjects(.ScreenobjectCount) = obj
        
        AddScreenObjectToLVZ = .ScreenobjectCount
        
        .ScreenobjectCount = .ScreenobjectCount + 1
    End With
End Function

Sub removeMapObjectFromLVZ(lvzIndex As Integer, objidx As Long)
'Removes a map object from an lvz
'@Pre: lvzIndex and index are within valid boundaries
'@Params
'    lvzIndex: index of lvz
'    index: index of map object
'@Post: map object is removed, mapobjectcount is decreased by 1

'    Dim i As Integer
'    For i = Index + 1 To lvz(lvzIndex).mapObjectCount - 1
'        lvz(lvzIndex).mapobjects(i - 1) = lvz(lvzIndex).mapobjects(i)
'    Next
    Dim objcount As Long
    
    With lvz(lvzIndex)
    
        objcount = .mapObjectCount - 1
        
        
        'Decrement the target indexes
        Dim i As Long
        For i = .mapobjects(objidx).layer To lyr_TopMost
            .nextIndexFor(i) = .nextIndexFor(i) - 1
        Next
        
        If objidx < objcount Then
            CopyMemory ByVal VarPtr(.mapobjects(objidx)), _
                        ByVal VarPtr(.mapobjects(objidx + 1)), _
                        (objcount - objidx) * LenB(.mapobjects(objidx))
        End If
        
        
        Call setMapObjectEmpty(.mapobjects(objcount))
        
        .mapObjectCount = objcount
        
        If .totalSelected > 0 Then
            
            For i = 0 To selectedObjectsCount - 1
                With selectedObjects(i)
                    If .lvzidx = lvzIndex Then
                        If .objidx > objidx Then .objidx = .objidx - 1
                    End If
                End With
            Next
        End If
        
    End With
    
    
'    frmGeneral.Label6.Caption = objcount & " objects remaining"
End Sub


 Function ListSelectedTrue() As String
    Dim i As Integer, j As Long
    ListSelectedTrue = ""
    
    For i = 0 To lvzcount - 1
        For j = 0 To lvz(i).mapObjectCount - 1
            If lvz(i).mapobjects(j).selected Then
                ListSelectedTrue = ListSelectedTrue & vbCrLf & i & "," & j
            End If
        Next
    Next
    
    ListSelectedTrue = ListSelectedTrue & vbCrLf & vbCrLf
    
    For j = 0 To selectedObjectsCount - 1
        ListSelectedTrue = ListSelectedTrue & vbCrLf & selectedObjects(j).lvzidx & "," & selectedObjects(j).objidx
    Next
End Function


Sub removeScreenObjectFromLVZ(lvzIndex As Integer, Index As Long)
'Removes a screen object from lvz
'@Pre: lvzIndex and index are within valid boundaries
'@Params
'    lvzIndex: index of lvz
'    index: index of screen object
'@Post: screen object is removed, screenobjectcount is decreased by 1

    Dim i As Long
    For i = Index + 1 To lvz(lvzIndex).ScreenobjectCount - 1
        lvz(lvzIndex).screenobjects(i - 1) = lvz(lvzIndex).screenobjects(i)
    Next
    
    Call setScreenObjectEmpty(lvz(lvzIndex).screenobjects(lvz(lvzIndex).ScreenobjectCount - 1))

    lvz(lvzIndex).ScreenobjectCount = lvz(lvzIndex).ScreenobjectCount - 1
    
    
End Sub

Function getIndexOfLVZ(name As String) As Long
'Get the index of a lvz with the specified name (not case sensitive)
'@Params name: name of lvz
'@Returns -1 if not found

    Dim i As Long
    For i = 0 To lvzcount - 1
        If LCase(lvz(i).name) = LCase(name) Then
            getIndexOfLVZ = i
            Exit Function
        End If
    Next
    
    getIndexOfLVZ = -1
End Function

Function getIndexOfFile(lvzIndex As Integer, path As String) As Long
'Retrieves the index of a file from a lvz with a given path (not case sensitive)
'Returns -1 if not found
    Dim i As Long
    For i = 0 To lvz(lvzIndex).filecount - 1
        If LCase(GetFileTitle(lvz(lvzIndex).files(i).path)) = LCase(GetFileTitle(path)) Then
            getIndexOfFile = i
            Exit Function
        End If
    Next
    
    getIndexOfFile = -1
End Function



Function getMapObjectImageID(lvzIndex As Integer, objectIndex As Long) As Long
    getMapObjectImageID = lvz(lvzIndex).mapobjects(objectIndex).imgidx
End Function

Function getScreenObjectImageID(lvzIndex As Integer, objectIndex As Long) As Long
    getScreenObjectImageID = lvz(lvzIndex).screenobjects(objectIndex).imgidx
End Function

Function SearchFile(filename As String, Optional filetype As LVZFileTypeEnum = lvz_other) As String
    Dim i As Long, j As Long
    Dim ext As String
    
    For i = 0 To lvzcount - 1
        For j = 0 To lvz(i).filecount - 1
            If filetype = lvz_image Then

                If LCase(GetFileNameWithoutExtension(lvz(i).files(j).path)) = LCase(filename) Then
                    ext = GetExtension(lvz(i).files(j).path)
                    
                    If IsImageType(ext) Then
                        SearchFile = lvz(i).files(j).path
                        Exit Function
                    End If
                End If
                
            ElseIf filetype = lvz_sound Then
            
                If LCase(GetFileNameWithoutExtension(lvz(i).files(j).path)) = LCase(filename) Then
                    ext = GetExtension(lvz(i).files(j).path)
                    
                    If ext = "wav" Or ext = "wa2" Then
                        SearchFile = lvz(i).files(j).path
                        Exit Function
                    End If
                End If
                
            Else
            
                If LCase(GetFileTitle(lvz(i).files(j).path)) = LCase(filename) Or _
                    LCase(lvz(i).files(j).path) = LCase(filename) Then
                    
                    SearchFile = lvz(i).files(j).path
                    Exit Function
                End If
                
            End If
        Next
    Next
    
    'no file found
    SearchFile = ""
End Function

Sub removeLVZ(Index As Integer)
    Dim i As Long
    For i = Index + 1 To lvzcount - 1
        lvz(i - 1) = lvz(i)
    Next
    
    Call setLVZEmpty(lvz(lvzcount - 1))
    
    lvzcount = lvzcount - 1
End Sub

Sub removeFileFromLVZ(lvzIndex As Integer, Index As Long)
    Dim i As Long
    For i = Index + 1 To lvz(lvzIndex).filecount - 1
        lvz(lvzIndex).files(i - 1) = lvz(lvzIndex).files(i)
    Next
    
    Call setFileEmpty(lvz(lvzIndex).files(lvz(lvzIndex).filecount - 1))

    lvz(lvzIndex).filecount = lvz(lvzIndex).filecount - 1
End Sub

Friend Function AddImageDefinitionToLVZ(Index As Integer, imgdef As LVZImageDefinition) As Integer

    If CountTotalImageDefinitions >= 256 Then
        AddImageDefinitionToLVZ = -1
    Else
    
        If lvz(Index).ImageDefinitionCount > UBound(lvz(Index).imagedefinitions) Then
            If (UBound(lvz(Index).imagedefinitions) + 5) >= 255 Then
                ReDim Preserve lvz(Index).imagedefinitions(255)
            Else
                ReDim Preserve lvz(Index).imagedefinitions(UBound(lvz(Index).imagedefinitions) + 5)
            End If
        End If
        
        lvz(Index).imagedefinitions(lvz(Index).ImageDefinitionCount) = imgdef
        
        AddImageDefinitionToLVZ = lvz(Index).ImageDefinitionCount
        
        lvz(Index).ImageDefinitionCount = lvz(Index).ImageDefinitionCount + 1
        
    End If
End Function

Sub removeImageDefinitionFromLVZ(lvzIndex As Integer, Index As Integer)
    Dim i As Long
    
    With lvz(lvzIndex)
        'Remove the definition from the array
        For i = Index + 1 To .ImageDefinitionCount - 1
            .imagedefinitions(i - 1) = .imagedefinitions(i)
        Next
        
        'Change the imgidx of mapobjects and screenobjects
        For i = 0 To .mapObjectCount - 1
            If .mapobjects(i).imgidx > Index Then
                .mapobjects(i).imgidx = .mapobjects(i).imgidx - 1
            End If
        Next
        
        For i = 0 To .ScreenobjectCount - 1
            If .screenobjects(i).imgidx > Index Then
                .screenobjects(i).imgidx = .screenobjects(i).imgidx - 1
            End If
        Next
        
        'Reset settings
        Call setImageDefinitionEmpty(.imagedefinitions(.ImageDefinitionCount - 1))
        
        'Decrease count
        .ImageDefinitionCount = .ImageDefinitionCount - 1
    
    End With
    
End Sub

Friend Sub setFileEmpty(File As LVZFileStruct)
    File.path = ""
    File.Type = lvz_other
End Sub

Friend Sub setMapObjectEmpty(obj As LVZMapObject)
    obj.displayTime = 0
    obj.layer = 0
    obj.imgidx = -1
    obj.mode = 0
    obj.objectID = 0
    obj.X = 0
    obj.Y = 0
    obj.selected = False

End Sub

Friend Sub setImageDefinitionEmpty(imgdef As LVZImageDefinition)
    imgdef.animationFramesX = 1 'default animation values
    imgdef.animationFramesY = 1
    imgdef.animationTime = 100
    imgdef.imagename = ""
    
    imgdef.CurrentFrame = 0
    imgdef.lastFrameChange = 0
End Sub

Friend Sub setScreenObjectEmpty(obj As LVZScreenObject)
    obj.imgidx = -1
    obj.displayTime = 0
    obj.layer = 0
    obj.mode = 0
    obj.objectID = 0
    obj.X = 0
    obj.Y = 0
    obj.typeX = scr_Normal
    obj.typeY = scr_Normal
End Sub

Friend Sub setLVZEmpty(lvz As LVZstruct)
    lvz.filecount = 0
    ReDim lvz.files(0)
    lvz.name = ""
End Sub

Friend Sub setMapObject(lvzIndex As Integer, mapobjIndex As Long, obj As LVZMapObject)
    lvz(lvzIndex).mapobjects(mapobjIndex) = obj
End Sub

Sub setMapObjectImageID(lvzIndex As Integer, mapobjIndex As Long, imgidx As Integer)
    lvz(lvzIndex).mapobjects(mapobjIndex).imgidx = imgidx
End Sub

Sub setScreenObjectImageID(lvzIndex As Integer, scrobjIndex As Long, imgidx As Integer)
    lvz(lvzIndex).screenobjects(scrobjIndex).imgidx = imgidx
End Sub

Sub setImageDefinitionFile(lvzIndex As Integer, imgdefIndex As Integer, filename As String)
    lvz(lvzIndex).imagedefinitions(imgdefIndex).imagename = filename
End Sub

Friend Sub setFile(lvzIndex As Integer, fileIndex As Long, File As LVZFileStruct)
    lvz(lvzIndex).files(fileIndex) = File
End Sub

Friend Sub setScreenObject(lvzIndex As Integer, scrobjIndex As Long, obj As LVZScreenObject)
    lvz(lvzIndex).screenobjects(scrobjIndex) = obj
End Sub

Friend Sub setImageDefinition(lvzIndex As Integer, imgdefIndex As Integer, imgdef As LVZImageDefinition)
    lvz(lvzIndex).imagedefinitions(imgdefIndex) = imgdef
End Sub

Friend Function getLVZFileType(path As String) As LVZFileTypeEnum
    Select Case GetExtension(GetFileTitle(path))
    
        Case "jpg", "bmp", "bm2", "gif", "png", "jpeg"
            getLVZFileType = lvz_image
        Case "wa2", "wav"
            getLVZFileType = lvz_sound
        Case Else
            getLVZFileType = lvz_other
            
    End Select
End Function

Private Sub buildSingleZoomLVZImage(picidx As Integer, picWidth As Integer, picHeight As Integer, zoom As Single)

'    Set picLVZImagesZoomPxl(picidx) = New clsDisplayLayer
'    Set picLVZImagesZoom1_4(picidx) = New clsDisplayLayer
'    Set picLVZImagesZoom1_2(picidx) = New clsDisplayLayer
'    Set picLVZImagesZoom2_1(picidx) = New clsDisplayLayer
    
    
    Dim srcDC As Long
    
    
    
    If zoom = 1 Then
        'Nothing to do, we'll be using picLVZImages(picidx) directly
        Set picLVZImagesZoom(picidx) = Nothing
    Else
        srcDC = picLVZImages(picidx).hDC
        Set picLVZImagesZoom(picidx) = New clsDisplayLayer
        
        With picLVZImagesZoom(picidx)
            SetStretchBltMode .hDC, COLORONCOLOR
            
            If zoom = 2 Then
                Call .Resize(picWidth * 2, picHeight * 2, False)
'                SetStretchBltMode .hDC, COLORONCOLOR
                
                StretchBlt .hDC, 0, 0, picWidth * 2, picHeight * 2, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
            ElseIf zoom = 1 / 2 Then
                Call .Resize(picWidth \ 2 + 1, picHeight \ 2 + 1, False)
'                SetStretchBltMode .hDC, HALFTONE
                
                StretchBlt .hDC, 0, 0, picWidth \ 2, picHeight \ 2, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
            ElseIf zoom = 1 / 4 Then
                Call .Resize(picWidth \ 4 + 3, picHeight \ 4 + 3, False)
'                SetStretchBltMode .hDC, HALFTONE
                
                StretchBlt .hDC, 0, 0, picWidth \ 4, picHeight \ 4, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
            Else
                'Pixel zoom 1:16
                Call .Resize(picWidth \ TILEW + 15, picHeight \ TILEH + 15, False)
'                SetStretchBltMode .hDC, COLORONCOLOR
                
                StretchBlt .hDC, 0, 0, picWidth \ TILEW, picHeight \ TILEH, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
            End If
        End With
    

    End If
    
'    With picLVZImagesZoomPxl(picidx)
'        Call .Resize(picWidth \ TILEW + 15, picHeight \ TILEH + 15, False)
'        SetStretchBltMode .hDC, COLORONCOLOR
'
'        StretchBlt .hDC, 0, 0, picWidth \ TILEW, picHeight \ TILEH, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
'    End With
'
'
'    With picLVZImagesZoom1_4(picidx)
'        Call .Resize(picWidth \ 4 + 3, picHeight \ 4 + 3, False)
'        SetStretchBltMode .hDC, HALFTONE
'
'        StretchBlt .hDC, 0, 0, picWidth \ 4, picHeight \ 4, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
'    End With
'
'
'    With picLVZImagesZoom1_2(picidx)
'        Call .Resize(picWidth \ 2 + 1, picHeight \ 2 + 1, False)
'        SetStretchBltMode .hDC, HALFTONE
'
'        StretchBlt .hDC, 0, 0, picWidth \ 2, picHeight \ 2, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
'    End With
'
'
'    With picLVZImagesZoom2_1(picidx)
'        Call .Resize(picWidth * 2, picHeight * 2, False)
'        SetStretchBltMode .hDC, HALFTONE
'
'        StretchBlt .hDC, 0, 0, picWidth * 2, picHeight * 2, srcDC, 0, 0, picWidth, picHeight, vbSrcCopy
'    End With
    

End Sub


Private Sub buildLibraryImages(lvzidx As Integer, imgidx As Integer)
    Dim wFactor As Single, hFactor As Single
    Dim frameX As Integer, frameY As Integer
    Dim previewWidth As Integer, previewHeight As Integer
    
    Dim srcDC As Long
    
    Dim X As Integer, Y As Integer
    Dim picWidth As Integer, picHeight As Integer
    
    X = 0
    Y = 0
    
    With lvz(lvzidx).imagedefinitions(imgidx)
        picWidth = .picWidth
        picHeight = .picHeight
        
        srcDC = picLVZImages(.picboxIdx).hDC
        
        Set picLVZImagesLibrary(.picboxIdx) = New clsDisplayLayer
        
        With picLVZImagesLibrary(.picboxIdx)
            Call .Resize(ImageTilesetW * lvz(lvzidx).imagedefinitions(imgidx).animationFramesX, _
                        ImageTilesetH * lvz(lvzidx).imagedefinitions(imgidx).animationFramesY, _
                        False)
                        
'            .width = ImageTilesetW * lvz(lvzidx).imagedefinitions(imgidx).animationFramesX
'            .height = ImageTilesetH * lvz(lvzidx).imagedefinitions(imgidx).animationFramesY
            
            For frameX = 0 To lvz(lvzidx).imagedefinitions(imgidx).animationFramesX - 1
                Y = 0
                For frameY = 0 To lvz(lvzidx).imagedefinitions(imgidx).animationFramesY - 1
                    
                    Call DrawImagePreviewCoords(srcDC, frameX * picWidth, frameY * picHeight, CLng(picWidth), CLng(picHeight), .hDC, CLng(X), CLng(Y), ImageTilesetW, ImageTilesetH, frmGeneral.TilesetBackgroundColor)
                    
                    
'                    If picWidth <= ImageTilesetW And picHeight <= ImageTilesetH Then
'                        'Image smaller than the frame, bitblt 1:1
'                        BitBlt .hDC, (ImageTilesetW - picWidth) \ 2 + X, (ImageTilesetH - picHeight) \ 2 + Y, picWidth, picHeight, srcDC, frameX * picWidth, frameY * picHeight, vbSrcCopy
'
'                    Else
'
'                        wFactor = picWidth / ImageTilesetW
'                        hFactor = picHeight / ImageTilesetH
'
'                        If hFactor > wFactor Then
'                            previewWidth = intMaximum(picWidth / hFactor, 1)
'                            previewHeight = ImageTilesetH
'
'
'                            SetStretchBltMode .hDC, IIf(hFactor >= 3, COLORONCOLOR, HALFTONE)
'                        Else
'                            previewWidth = ImageTilesetW
'                            previewHeight = intMaximum(picHeight / wFactor, 1)
'
'                            SetStretchBltMode .hDC, IIf(wFactor >= 3, COLORONCOLOR, HALFTONE)
'                        End If
'
'                        StretchBlt .hDC, (ImageTilesetW - previewWidth) \ 2 + X, (ImageTilesetH - previewHeight) \ 2 + Y, previewWidth, previewHeight, srcDC, frameX * picWidth, frameY * picHeight, picWidth, picHeight, vbSrcCopy
'                    End If
                    
                    Y = Y + ImageTilesetH
                Next
                
                X = X + ImageTilesetW
            Next
        End With
    End With
End Sub

Private Sub buildSingleLVZImage(picboxIdx As Integer, lvzidx As Integer, imgidx As Integer)
    Dim fullwidth As Integer, fullheight As Integer

    Set picLVZImages(picboxIdx) = New clsDisplayLayer
    
    picUsed(picboxIdx) = True
    
    With lvz(lvzidx).imagedefinitions(imgidx)
    
        Call picLVZImages(picboxIdx).LoadImageFile(SearchFile(.imagename), frmGeneral.pictemp, fullwidth, fullheight)
        
        .picboxIdx = picboxIdx
        
        .picWidth = fullwidth \ .animationFramesX
        .picHeight = fullheight \ .animationFramesY
        
        Call buildSingleZoomLVZImage(picboxIdx, .picWidth, .picHeight, parent.magnifier.zoom)
        
    End With

    Call buildLibraryImages(lvzidx, imgidx)
End Sub

Sub buildAllZoomImages(zoom As Single)
    Dim i As Integer, j As Integer

        
    For i = 0 To lvzcount - 1
        For j = 0 To lvz(i).ImageDefinitionCount - 1
            With lvz(i).imagedefinitions(j)
                
                Call buildSingleZoomLVZImage(.picboxIdx, .picWidth, .picHeight, zoom)

            End With
        Next
        
    Next
End Sub

Sub buildAllLVZImages()
    'call this to load all lvz images into their picture boxes
    'Only call this if the structure of the lvz won't change.
    'If it changes, you muse recall this method!, else the wrong or removed lvz
    'files will be assigned to the images
    Dim i As Integer, j As Integer
    
    
    
    Dim imgCount As Integer
    imgCount = 0
    
    Call ResizePicLVZArrays(CountTotalImageDefinitions, False)
    
    For i = 0 To lvzcount - 1
        For j = 0 To lvz(i).ImageDefinitionCount - 1
            
            With lvz(i).imagedefinitions(j)
                
                If SearchFile(.imagename) <> "" Then
                    
                    Call buildSingleLVZImage(imgCount, i, j)

                    imgCount = imgCount + 1
                Else
                    AddDebug ("Error loading LVZ image, file '" & lvz(i).imagedefinitions(j).imagename & "' not found in any lvz")
                End If
            
            End With
        Next
        
    Next
    
    Call frmGeneral.cTileset.DrawLVZTileset(True)
'    Call parent.tileset.DrawLVZTileset(True)
End Sub

'Private Sub tryLoadingPicLVZ(idx As Integer)
'    On Error Resume Next
'    Load parent.picLVZItem(idx)
'    Load parent.picLVZItemZoom1_2(idx)
'    Load parent.picLVZItemZoom1_4(idx)
'    Load parent.picLVZItemZoom2(idx)
'    Load parent.picLVZItemZoomPix(idx)
'    Load parent.picLVZLibrary(idx)
'
'    parent.picLVZItem(idx).AutoRedraw = True
'    parent.picLVZItemZoom1_2(idx).AutoRedraw = True
'    parent.picLVZItemZoom1_4(idx).AutoRedraw = True
'    parent.picLVZItemZoom2(idx).AutoRedraw = True
'    parent.picLVZItemZoomPix(idx).AutoRedraw = True
'    parent.picLVZLibrary(idx).AutoRedraw = True
'End Sub

Sub exportINI(path As String, lvzidx As Integer)
    ' path is where lvl file is ?
    ' anyhow, this allows more flexibility where to save them
    
'    If Mid$(path, Len(path) - 1, 1) <> "\" Then
'        lvzpath = path & "\" & LVZ(lvzIdx).name
'    Else
'        lvzpath = path & LVZ(lvzIdx).name
'    End If
    If GetExtension(path) <> "ini" Then
        path = GetPathTo(path) & GetFileNameWithoutExtension(path) & ".ini"
    End If
    
'    If Not CheckOverwrite(path) Then Exit Sub
    
    
    Dim f As Integer

    Dim i As Long
    
    f = FreeFile
    Open path For Output As #f
    
    With lvz(lvzidx)
    
        'HEADER
        
        Print #f, "Outfile=" & .name
        
        Print #f, ""
        
        
        'FILES
        For i = 0 To .filecount - 1
            Print #f, "File=" & GetFileTitle(.files(i).path)
        Next
        
        Print #f, ""
        
        'IMAGE DEFINITIONS
        Print #f, "[objectimages]"
        
        For i = 0 To .ImageDefinitionCount - 1
            With .imagedefinitions(i)
                Print #f, "IMAGE" & i & "=" & .imagename & _
                                            "," & .animationFramesX & _
                                            "," & .animationFramesY & _
                                            "," & .animationTime
            End With
        Next
        
        Print #f, ""
        
        
        'MAP OBJECTS
        Print #f, "[mapobjects]"
        
        For i = 0 To .mapObjectCount - 1
            With .mapobjects(i)
                Print #f, .X & "," & .Y & ",IMAGE" & .imgidx & "," & LVZLayerName(.layer) & "," & LVZModeName(.mode) & "," & .displayTime & "," & .objectID
            End With
        Next i
        
        Print #f, ""
        
        'SCREEN OBJECTS
        Print #f, "[screenobjects]"
        
        For i = 0 To .ScreenobjectCount - 1
            With .screenobjects(i)
                Print #f, LVZScreenObjectPrefix(.typeX) & .X & "," & LVZScreenObjectPrefix(.typeY) & .Y & ",IMAGE" & .imgidx & "," & LVZLayerName(.layer) & "," & LVZModeName(.mode) & "," & .displayTime & "," & .objectID
            End With
        Next i
    
    End With
    
    Close #f

    If FileExists(path) Then
        ShellExecute 0&, vbNullString, "notepad", path, vbNullString, vbNormalFocus
    End If
End Sub


Sub exportLVZ(path As String, lvzidx As Integer)
    ' path is where lvl file is ?
    ' anyhow, this allows more flexibility where to save them
    Dim lvzpath As String
    
'    If Mid$(path, Len(path) - 1, 1) <> "\" Then
'        lvzpath = path & "\" & LVZ(lvzIdx).name
'    Else
'        lvzpath = path & LVZ(lvzIdx).name
'    End If
    If GetExtension(path) <> "lvz" Then
        lvzpath = GetPathTo(path) & GetFileNameWithoutExtension(path) & ".lvz"
    Else
        lvzpath = path
    End If

    lvz(lvzidx).path = path
    
    Dim f As Integer
    Dim f2 As Integer

    Dim header As LVZheader
    Dim i As Long, j As Long
    Dim zlib As New zlib
    Dim Data() As Byte
    Dim DecompressedSize As Long
    Dim zlibret As Long
    
    'does file exists?
    
    If Not CheckOverwrite(lvzpath) Then Exit Sub
'    If FileExists(lvzpath) Then
'        If confirmOverwrite Then
'            If MessageBox(lvzpath & " already exists. Do you want to overwrite it?", vbYesNo + vbQuestion, "Confirm overwrite") = vbYes Then
'                If FileExists(lvzpath) Then Kill lvzpath
'            Else
'                Exit Sub
'            End If
'        Else
'            Kill lvzpath
'        End If
'    End If
    
    f = FreeFile
    Open lvzpath For Binary As #f

    Put #f, , stringToBytes("CONT")
    If lvz(lvzidx).mapObjectCount > 0 Or lvz(lvzidx).ScreenobjectCount > 0 Or lvz(lvzidx).ImageDefinitionCount > 0 Then
        'An additionnal section will be needed for object definitions
        Put #f, , LongtoBytes(CLng(lvz(lvzidx).filecount + 1))
    Else
        Put #f, , LongtoBytes(CLng(lvz(lvzidx).filecount))
    End If
    
    'put object definitions section
    If lvz(lvzidx).mapObjectCount > 0 Or lvz(lvzidx).ScreenobjectCount > 0 Or lvz(lvzidx).ImageDefinitionCount > 0 Then
        'We still have to compress this data, so build a temporary file
        Dim tmpobjpath As String
        tmpobjpath = GetPathTo(lvzpath) & "\" & Format(Now, "yyddhhmmss_") & lvz(lvzidx).name & "_tempobjdefs.clv"
        
        If FileExists(tmpobjpath) Then Kill tmpobjpath
        
        f2 = FreeFile
        Open tmpobjpath For Binary As #f2
        
        'Object section header
        Put #f2, , stringToBytes("CLV2")
        
        'Object count
        Put #f2, , LongtoBytes(CLng(lvz(lvzidx).mapObjectCount + lvz(lvzidx).ScreenobjectCount))
        
        'Image definitions count
        Put #f2, , LongtoBytes(CLng(lvz(lvzidx).ImageDefinitionCount))
        
        'Put map objects data
        For i = 0 To lvz(lvzidx).mapObjectCount - 1
            'IsMapObject + ObjectID (u1 + u15)
            Put #f2, , IntegerToBytes(CLng(lvz(lvzidx).mapobjects(i).objectID * 2& + 1))

            'X/Y (i16)
            Put #f2, , lvz(lvzidx).mapobjects(i).X
            Put #f2, , lvz(lvzidx).mapobjects(i).Y
            
            'Image Number (u8)
            Put #f2, , CByte(lvz(lvzidx).mapobjects(i).imgidx)
            
            'Layer (u8)
            Put #f2, , CByte(lvz(lvzidx).mapobjects(i).layer)
            
            'Display time (u12) + Display mode (u4)
            Put #f2, , CByte(lvz(lvzidx).mapobjects(i).displayTime Mod 256)
            Put #f2, , CByte((lvz(lvzidx).mapobjects(i).displayTime \ 256) + lvz(lvzidx).mapobjects(i).mode * 16)
        Next
        
        'Put screen objects data
        For i = 0 To lvz(lvzidx).ScreenobjectCount - 1
            'IsMapObject + ObjectID (u1 + u15)
            Put #f2, , IntegerToBytes(CLng(lvz(lvzidx).screenobjects(i).objectID * 2&))

            Dim b(1) As Byte
            Put #f2, , CByte(lvz(lvzidx).screenobjects(i).typeX + (SignedTo12bits(lvz(lvzidx).screenobjects(i).X) Mod 16) * 16)
            Put #f2, , CByte(SignedTo12bits(lvz(lvzidx).screenobjects(i).X) \ 16)
            
            'Y type (u4) + Y (i12)
            Put #f2, , CByte(lvz(lvzidx).screenobjects(i).typeY + (SignedTo12bits(lvz(lvzidx).screenobjects(i).Y) Mod 16) * 16)
            Put #f2, , CByte(SignedTo12bits(lvz(lvzidx).screenobjects(i).Y) \ 16)
            
            'Image Number (u8)
            Put #f2, , CByte(lvz(lvzidx).screenobjects(i).imgidx)
            
            'Layer (u8)
            Put #f2, , CByte(lvz(lvzidx).screenobjects(i).layer)
            
            'Display time (u12) + Display mode (u4)
            Put #f2, , CByte(lvz(lvzidx).screenobjects(i).displayTime Mod 256)
            Put #f2, , CByte((lvz(lvzidx).screenobjects(i).displayTime \ 256) + lvz(lvzidx).screenobjects(i).mode * 16)
        Next
        
        'Put image definitions
        For i = 0 To lvz(lvzidx).ImageDefinitionCount - 1
            'Frames X
            Put #f2, , lvz(lvzidx).imagedefinitions(i).animationFramesX
            
            'Frames Y
            Put #f2, , lvz(lvzidx).imagedefinitions(i).animationFramesY
            
            'Frame display time
            Put #f2, , lvz(lvzidx).imagedefinitions(i).animationTime
            
            'FileName (null-ended string)

            Put #f2, , stringToBytes(lvz(lvzidx).imagedefinitions(i).imagename)
            Put #f2, , CByte(0)
        Next
        
        
        Close #f2
        
        'Grab all the data we just put
        DecompressedSize = FileLen(tmpobjpath)
        ReDim Data(DecompressedSize - 1)
        
        Open tmpobjpath For Binary As #f2
        Get #f2, , Data
        Close #f2
        
        If FileExists(tmpobjpath) Then Kill tmpobjpath
        
        'Compress the data, and dump it in lvz file with its header
        zlibret = zlib.CompressData(Data)
        If zlibret <> 0 Then
            'error occured
            MessageBox "Zlib error " & zlibret & ". Could not compress the data for object definitions.", vbExclamation + vbOKOnly
        Else
                    'Header type
            Put #f, , stringToBytes("CONT")
            
            'Decompressed size
            Put #f, , LongtoBytes(DecompressedSize)
            
            'Filetime (=0 for object definitions)
            Put #f, , CLng(0)
            
            'Compressed Size
            Put #f, , LongtoBytes(UBound(Data) + 1)
            
            'File name ("" for object definitions)
            Put #f, , CByte(0) 'Null-ended string
            
            'Compressed Obj Definitions Data
            Put #f, , Data
        End If
    End If
    
    'put compressed files
    For i = 0 To lvz(lvzidx).filecount - 1
        
        'Grab data from file
        ReDim Data(FileLen(lvz(lvzidx).files(i).path) - 1)
        
        f2 = FreeFile
        Open lvz(lvzidx).files(i).path For Binary As #f2
        Get #f2, , Data
        Close #f2
        
        zlibret = zlib.CompressData(Data)
        If zlibret <> 0 Then
            'error occured
            MessageBox "Zlib error " & zlibret & ". Could not compress the data for file " & i & " (" & lvz(lvzidx).files(i).path & ").", vbExclamation + vbOKOnly
        Else
            'go on
        
            'Header type
            Put #f, , stringToBytes("CONT")
            
            'Decompressed size
            DecompressedSize = FileLen(lvz(lvzidx).files(i).path)
            Put #f, , LongtoBytes(DecompressedSize)
            
            'Filetime
            Put #f, , DoubletoBytes(CDbl(DateDiff("s", "1/1/1970 00:00:00", Format(Date, "d/m/yyyy") & " " & Format(time, "hh:mm:ss"))))
            
            'Compressed Size
            Put #f, , LongtoBytes(UBound(Data) + 1)
            
            'File name
            Put #f, , stringToBytes(GetFileTitle(lvz(lvzidx).files(i).path))
            Put #f, , CByte(0) 'Null-ended string
            
            'Compressed Data
            Put #f, , Data
            
        End If
        
        
    Next
         
    Close #f
    
    If FileLen(lvzpath) >= 4194304 Then
        MessageBox "Warning: " & GetFileTitle(lvzpath) & " is larger than 4MB, Continuum might not be able to use that file. Remove some files or split its content in several lvz files to reduce its size to make it usable.", vbOKOnly + vbExclamation
    End If

End Sub

Sub importLVZ(path As String, Optional rebuildImages As Boolean = True)
    '// DO STUFF TO ADD STUFF HERE
    Dim f As Integer
    Dim f2 As Integer
    
    Dim lvzpath As String
    Dim lvzname As String
    Dim lvzID As Integer
    
    Dim filename As String
    Dim DestPath As String
    
    Dim header As LVZheader
    Dim FILETIME(3) As Byte 'u32
    Dim i As Long, j As Long
    Dim zlib As New zlib
    
    f = FreeFile
    
    lvzpath = path
    lvzname = GetFileTitle(lvzpath)
    
    
    'does file exists?
    If Not FileExists(lvzpath) Then Exit Sub
    
    DestPath = App.path & "\DCME LVZ files\" & GetFileNameWithoutExtension(lvzpath)
    
    Open lvzpath For Binary As #f
    
    Get #f, , header
    
    If bytesToString(header.filetype) <> "CONT" Then
        'invalid LVZ
        MessageBox "This file is not a valid lvz. Expected header: 'CONT'. Found '" & bytesToString(header.filetype) & "'.", vbOKOnly + vbExclamation
    Else
        'Add lvz file
        i = 1
        While (Not AddLVZ(lvzname))
                'change the name until it works
                lvzname = GetFileTitle(lvzpath) & "(" & i & ")"
                i = i + 1
        Wend
        
        lvzID = getIndexOfLVZ(lvzname)
        
        Dim SectionCount As Long
        SectionCount = bytesToLong(header.size)
        
        For i = 1 To SectionCount
            'there are sections
            Get #f, , header
            
            If bytesToString(header.filetype) <> "CONT" Then
                'Bad compression header, do not use
                'Stop
                MessageBox "Section " & i & " with size " & bytesToLong(header.size) & " has bad compression header. Expected header: 'CONT'. Found '" & bytesToString(header.filetype) & "'.", vbOKOnly + vbExclamation
            Else
                Dim DecompressedSize As Double 'u32
                Dim CompressedSize As Double 'u32
                Dim CompressedData() As Byte
                
                Dim u32(3) As Byte
                Dim b As Byte
                
                'Section Decompressed Size
                DecompressedSize = bytesToDouble(header.size)
                'File Time '?
                Get #f, , FILETIME
                
                'Section Compressed Size
                Get #f, , u32
                CompressedSize = bytesToDouble(u32)
                
                'FileName (null terminated) , can be null if FileTime is 0
                filename = ""
                Do
                    Get #f, , b
                    If b <> 0 Then
                        filename = filename & Chr$(b)
                    End If
                Loop While (b <> 0)
                
                'Compressed Data
                ReDim CompressedData(CompressedSize - 1)
                Get #f, , CompressedData
                
                'Decompress the data if needed
                Dim ret As Long
                
                If CompressedSize <> DecompressedSize Then
                    ret = (zlib.DecompressData(CompressedData, CLng(DecompressedSize)))
                Else
                    ret = 0
                End If
                
                If ret <> 0 Then
                    'error occured
                    MessageBox "Error decompressing " & filename, vbOKOnly + vbExclamation
                Else
                    'Check if it is a file, or objects definitions
                    If filename = "" And bytesToDouble(FILETIME) = 0 Then
                        'OBJECT DEFINITIONS
                        'save definitions to file for now
                        Dim clv_type As ContinuumLevelObjectTypes
                        Dim clv_header As CLVheader
                        f2 = FreeFile
'450                           DestPath = App.path & "\DCME LVZ files\" & GetFileNameWithoutExtension(lvzpath)
                        CreateDir (DestPath)
                        Open DestPath & "\" & "ObjDef.clv" For Binary As #f2
                        Put #f2, , CompressedData
                        
                        Seek #f2, 1
                        
                        'Start reading CLV data
                        Get #f2, , clv_header
                        
                        If bytesToString(clv_header.filetype) = "CLV1" Then
                            clv_type = CLV1
                        ElseIf bytesToString(clv_header.filetype) = "CLV2" Then
                            clv_type = CLV2
                        Else
                            clv_type = Invalid 'nothing to do with that crap
                        End If
                        
                        If clv_type = CLV1 Or clv_type = CLV2 Then
                            Dim objcount As Double
                            Dim imgCount As Double
                            Dim clv_bytes(0 To 9) As Byte
                            objcount = bytesToDouble(clv_header.objcount)
                            imgCount = bytesToDouble(clv_header.imgCount)
                            
                            For j = 1 To objcount
                            'Loop through objects (map object and screenobjects)
                            
                                'grab 10 bytes of data
                                Get #f2, , clv_bytes
                                
                                Dim CLV As New bitArray
                                Call CLV.SetData(clv_bytes)

                                
                                'consider it is a mapobject for now, we'll store it later
                                If CLV.GetBit(0) Then
                                    'This is a map object
                                    Dim tmpMapObj As LVZMapObject
                                    tmpMapObj.objectID = CLV.GetValue(1, 15)
                                    tmpMapObj.X = UnsignedToInteger(CLng(CLV.GetValue(16, 16)))
                                    tmpMapObj.Y = UnsignedToInteger(CLng(CLV.GetValue(32, 16)))
                                    tmpMapObj.imgidx = CLV.GetValue(48, 8)
                                    tmpMapObj.layer = CLV.GetValue(56, 8)
                                    tmpMapObj.displayTime = CLV.GetValue(64, 12)
                                    tmpMapObj.mode = CLV.GetValue(76, 4)
                                    
                                    If AddMapObjectToLVZ(lvzID, tmpMapObj) <> -1 Then
                                        'Map object added
                                    Else
                                        MessageBox "Error adding mapobject to lvz", vbCritical + vbOKOnly
                                    End If
                                    
                                Else
                                    'This is a screen object
                                    Dim tmpScrObj As LVZScreenObject
                                    If clv_type = CLV1 Then
                                        tmpScrObj.X = CLng(CLV.GetValue(16, 16))
                                        If tmpScrObj.X > 2047 Then tmpScrObj.X = tmpScrObj.X - 4096
                                        tmpScrObj.Y = CLng(CLV.GetValue(32, 16))
                                        If tmpScrObj.Y > 2047 Then tmpScrObj.Y = tmpScrObj.Y - 4096
                                        tmpScrObj.typeX = scr_Normal
                                        tmpScrObj.typeY = scr_Normal
                                    Else
                                        tmpScrObj.typeX = CInt(CLV.GetValue(16, 4))
                                        tmpScrObj.X = SignedFrom12bits(CInt(CLV.GetValue(20, 12)))
                                        tmpScrObj.typeY = CInt(CLV.GetValue(32, 4))
                                        tmpScrObj.Y = SignedFrom12bits(CInt(CLV.GetValue(36, 12)))
                                    End If
                                    
                                    'These are exactly the same as map objects
                                    tmpScrObj.objectID = CLV.GetValue(1, 15)

                                    tmpScrObj.imgidx = CLV.GetValue(48, 8)
                                    tmpScrObj.layer = CLV.GetValue(56, 8)
                                    tmpScrObj.displayTime = CLV.GetValue(64, 12)
                                    tmpScrObj.mode = CLV.GetValue(76, 4)
                                    
                                    If AddScreenObjectToLVZ(lvzID, tmpScrObj) <> -1 Then
                                        'Map object added
                                    Else
                                        MessageBox "Error adding screenobject to lvz", vbCritical + vbOKOnly
                                    End If
                                End If

                            Next
                            'Done with object count
                            
                            'Now for image definitions
                            
                            For j = 1 To imgCount
                                Dim tmpImgDef As LVZImageDefinition
                                Dim tmpInt As Integer
                                
                                Get #f2, , tmpImgDef.animationFramesX
                                Get #f2, , tmpImgDef.animationFramesY
                                Get #f2, , tmpImgDef.animationTime
'                                tmpImgDef.animationTime = IntegerToUnsigned(tmpInt)
                                
                                If tmpImgDef.animationTime <= 0 Then tmpImgDef.animationTime = 100
                                
                                tmpImgDef.imagename = ""
                                Dim tmpByte As Byte
                                Do
                                    Get #f2, , tmpByte
                                    If tmpByte <> 0 Then tmpImgDef.imagename = tmpImgDef.imagename & Chr$(tmpByte)
                                Loop While tmpByte <> 0
                            
                                If AddImageDefinitionToLVZ(lvzID, tmpImgDef) <> -1 Then
                                    'Added image def successfully
                                Else
                                    MessageBox "Error adding image definition", vbCritical + vbOKOnly
                                End If
                            Next
                            'Done with image definitions count
                            
                        Else
                            MessageBox "Object definitions header was invalid. Your files will be loaded correctly, but all image definitions, map objects and screen objects will be missing. Expected header: 'CLV1' or 'CLV2'. Found '" & bytesToString(clv_header.filetype) & "'.", vbOKOnly + vbExclamation
                        End If
                        
                        Close #f2
                        'Done with CLV1/2 object definitions
                        
                        'Delete the file since everything has been loaded
                        Kill DestPath & "\" & "ObjDef.clv"
                        
                    Else
                        'This is a normal file (could be anything...)
                        
                        If filename = "" Then filename = GetFileNameWithoutExtension(lvzpath) & "_file" & i
                        
                        'Save file
                        f2 = FreeFile
'1250                          DestPath = App.path & "\DCME LVZ files\" & GetFileNameWithoutExtension(lvzpath)
                        CreateDir (DestPath)
                        Open DestPath & "\" & filename For Binary As #f2
                        
                        Put #f2, , CompressedData
                        
                        Close #f2
                        
                        If AddFileToLVZ(lvzID, DestPath & "\" & filename) <> -1 Then
                            'file added
                        Else
'                            If MessageBox("Error adding file " & filename & " to lvz. The file already exists in the same lvz.", vbExclamation + vbOKCancel) = vbCancel Then
'                                Close #f
'                                Exit Sub
'                            End If
                                
                        End If
                            
                        
                    End If
                End If
                
            End If
        Next
        
    End If
    
    Close #f
    
    ' in case you have a loop that imports all stored lvz in elvl (that you don't rebuild the images everytime)
    If rebuildImages Then
        Call buildAllLVZImages
    End If
End Sub

Function getImageWidth(lvzidx As Integer, imgidx As Integer) As Integer

    getImageWidth = lvz(lvzidx).imagedefinitions(imgidx).picWidth
    
    'parent.picLVZItem(lvz(lvzidx).imagedefinitions(imgidx).picboxIdx).width \ lvz(lvzidx).imagedefinitions(imgidx).animationFramesX
End Function

Function getImageHeight(lvzidx As Integer, imgidx As Integer) As Integer
    getImageHeight = lvz(lvzidx).imagedefinitions(imgidx).picHeight
    
'    getImageHeight = parent.picLVZItem(lvz(lvzidx).imagedefinitions(imgidx).picboxIdx).height \ lvz(lvzidx).imagedefinitions(imgidx).animationFramesY
End Function



Function GetMapObjectWidth(lvzidx As Integer, objidx As Long) As Integer
    With lvz(lvzidx)
        GetMapObjectWidth = .imagedefinitions((.mapobjects(objidx).imgidx)).picWidth
    End With
'    = getImageWidth(lvzidx, lvz(lvzidx).mapobjects(objIdx).imgidx)
End Function

Function GetMapObjectHeight(lvzidx As Integer, objidx As Long) As Integer
    With lvz(lvzidx)
        GetMapObjectHeight = .imagedefinitions((.mapobjects(objidx).imgidx)).picHeight
    End With
'    GetMapObjectHeight = getImageHeight(lvzidx, lvz(lvzidx).mapobjects(objIdx).imgidx)
End Function

Function GetScreenObjectWidth(lvzidx As Integer, objidx As Long) As Integer
    With lvz(lvzidx)
        GetScreenObjectWidth = .imagedefinitions((.screenobjects(objidx).imgidx)).picWidth
    End With
'    GetScreenObjectWidth = getImageWidth(lvzidx, lvz(lvzidx).screenobjects(objIdx).imgidx)
End Function

Function GetScreenObjectHeight(lvzidx As Integer, objidx As Long) As Integer
    With lvz(lvzidx)
        GetScreenObjectHeight = .imagedefinitions((.screenobjects(objidx).imgidx)).picHeight
    End With
'    GetScreenObjectHeight = getImageHeight(lvzidx, lvz(lvzidx).screenobjects(objIdx).imgidx)
End Function


Private Function ObjectIsInArea(lvzidx As Integer, objidx As Long, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, partial As Boolean) As Boolean
    'Returns true if the specified mapobject is in the specified area
    
    'lvzidx, objidx       : mapobject to check
    'Left,Top,Right,Bottom: area to check
    'partial              : If true, function will return true if at least
    '                                1 pixel of the object is within the area
    '                       If false, function will only return true if
    '                                 the entire object is within the area
    'avoidBlack           : If true, function will try to find a visible pixel of the image
    '                       This will be used when selecting an object where many objects
    '                       are present at the same coordinate
    '                       This is only valid if 'partial' is set to true
    
    With lvz(lvzidx).mapobjects(objidx)

        If partial Then
            If .X <= Right And .Y <= Bottom Then
                If .X + GetMapObjectWidth(lvzidx, objidx) >= Left Then
                    ObjectIsInArea = (.Y + GetMapObjectHeight(lvzidx, objidx) >= Top)
                End If
            End If
        Else
            If .X >= Left And .Y >= Top Then
                If .X + GetMapObjectWidth(lvzidx, objidx) <= Right Then
                    ObjectIsInArea = (.Y + GetMapObjectHeight(lvzidx, objidx) <= Bottom)
                End If
            End If
        End If
        
    End With
End Function





Friend Function getMapObjectsInArea(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, _
                             partial As Boolean, ByRef mapobjs() As MapObjectRef) As Long
    'Puts all objects in the given area in the mapobjs() array
    'Returns the number of objects returned
    Dim count As Long
    Dim i As Integer, j As Long
    
    For i = 0 To lvzcount - 1
        For j = 0 To lvz(i).mapObjectCount - 1
            If ObjectIsInArea(i, j, Left, Top, Right, Bottom, partial) Then
                If count = 0 Then
                    ReDim mapobjs(64)
                ElseIf count > UBound(mapobjs) Then
                    ReDim Preserve mapobjs(count + 64)
                End If
                    
                mapobjs(count).lvzidx = i
                mapobjs(count).objidx = j
                
                count = count + 1
            End If
        Next
    Next
    
    If count > 0 Then
        ReDim Preserve mapobjs(count - 1)
    End If
    getMapObjectsInArea = count
End Function

Friend Function getMapObjectAtPos(ByVal X As Integer, ByVal Y As Integer, ByRef lvzidx As Integer, ByRef objidx As Long) As Long
    'Gets the topmost object at X,Y
    
    'X, Y          : Coordinate to check
    'lvzidx, objidx: Return values of the mapobject
    
    getMapObjectAtPos = -1
    
    Dim i As Integer, j As Long
    
    'We have to check them in the reverse order they are drawn (top to bottom)
    For i = lvzcount - 1 To 0 Step -1

        For j = lvz(i).mapObjectCount - 1 To 0 Step -1
            
            If ObjectIsInArea(i, j, X, Y, X, Y, True) Then
                'Found one!
                lvzidx = i
                objidx = j
                
                getMapObjectAtPos = j
                
                'If it's a black pixel, keep searching, we might find other visible images under it
                With lvz(i).mapobjects(j)
                    If GetPixel(picLVZImages(lvz(i).imagedefinitions(.imgidx).picboxIdx).hDC, X - .X, Y - .Y) Then
                        Exit Function
                    End If
                End With
            End If

        Next
        
    Next

    
End Function

Function HasSelection() As Boolean
    HasSelection = selectedObjectsCount
End Function


Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim zoom As Single
    zoom = parent.magnifier.zoom


    
    firstX = (parent.Hscr.value + X) / zoom
    firstY = (parent.Vscr.value + Y) / zoom
    lastX = 0
    lastY = 0
    
    Dim lvzidx As Integer, objidx As Long
    If getMapObjectAtPos(firstX, firstY, lvzidx, objidx) <> -1 Then
        If Button = vbLeftButton Then
            If IsShift(Shift) Or IsControl(Shift) Then
                'Start to add or remove a selection
                State = S_SelectingSingle
                
                parent.picPreview.MousePointer = MousePointerConstants.vbCrosshair
            Else
                If Not lvz(lvzidx).mapobjects(objidx).selected Then
                'If not selected, select it and start moving
                    Call EmptySelection
                    Call AddToSelection(lvzidx, objidx)
                    
                'Else, just start moving
                End If
                
                State = S_Moving
                
                parent.picPreview.MousePointer = MousePointerConstants.vbSizeAll
            End If
            
        
        ElseIf Button = vbRightButton Then
        
        End If
    Else
        If Button = vbLeftButton Then
            'Unselect all if not pressing shift or ctrl
            If Shift = 0 Then Call EmptySelection
        
            State = S_SelectingSingle
            
            parent.picPreview.MousePointer = MousePointerConstants.vbCrosshair
            
        ElseIf Button = vbRightButton Then
        
        End If
    End If

    Call parent.UpdatePreview
End Sub



Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'    If bDEBUG Then
'        frmGeneral.Label6.Caption = ((parent.hScr.value * (TILEW / parent.currenttilew) + X) \ parent.currenttilew) * parent.currenttilew & " - " & ((parent.vScr.value * (TILEW / parent.currenttilew) + Y) \ parent.currenttilew) * parent.currenttilew
'    End If

    Dim curX As Integer, curY As Integer
    Dim zoom As Single
    zoom = parent.magnifier.zoom
        
        
    Dim hoverobjs() As MapObjectRef
    
    
    curX = (parent.Hscr.value + X) / zoom
    curY = (parent.Vscr.value + Y) / zoom
        
    If State = S_Moving Then

        Dim newX As Integer, newY As Integer
        newX = ((curX - firstX) \ c_Snap) * c_Snap
        newY = ((curY - firstY) \ c_Snap) * c_Snap
        
        If newX <> lastX Or newY <> lastY Then
            Call MoveSelection(newX - lastX, newY - lastY)
            
            lastX = newX
            lastY = newY
        
            Call parent.RedrawLvzUnderLayer(False)
            Call parent.RedrawLvzOverLayer(True)
        End If
    
    ElseIf State = S_SelectingSingle Or State = S_SelectingArea Then
        'We moved mouse after clicking for selection
        State = S_SelectingArea

        'Get all objects in that area
        Dim selarea As area
        
        selarea.Left = intMinimum(firstX, curX)
        selarea.Top = intMinimum(firstY, curY)
        selarea.Right = intMaximum(firstX, curX)
        selarea.Bottom = intMaximum(firstY, curY)
        
        Dim count As Long
        
        With selarea
            count = getMapObjectsInArea(.Left, .Top, .Right, .Bottom, curX < firstX, hoverobjs)
        End With
        
        Call parent.UpdatePreview(False, False)
        
        'Highlight them
        If count > 0 Then Call HighLightObjects(hoverobjs, count, vbWhite)
            
        Call DrawRectangle(parent.picPreview.hDC, _
                            selarea.Left * zoom - parent.Hscr.value, _
                            selarea.Top * zoom - parent.Vscr.value, _
                            selarea.Right * zoom - parent.Hscr.value, _
                            selarea.Bottom * zoom - parent.Vscr.value, _
                            vbCyan)
        
'        With parent.shpLVZsel
'            .Left = intMinimum(firstX, curX) * zoom - parent.hScr.value
'            .Top = intMinimum(firstY, curY) * zoom - parent.vScr.value
'            .width = Abs(curX - firstX) * zoom
'            .height = Abs(curY - firstY) * zoom
'            .visible = True
'        End With
        
        parent.picPreview.Refresh
    Else
        'Just moving the mouse around
        ReDim hoverobjs(0)
        
        Call parent.UpdatePreview(False, False)
        
        If getMapObjectAtPos(curX, curY, hoverobjs(0).lvzidx, hoverobjs(0).objidx) <> -1 Then
            
            Call HighLightObjects(hoverobjs, 1, IIf(lvz(hoverobjs(0).lvzidx).mapobjects(hoverobjs(0).objidx).selected, vbBlue, parent.cursorcolor))
            
            If IsControl(Shift) Then
                parent.picPreview.MousePointer = MousePointerConstants.vbDefault
            Else
                parent.picPreview.MousePointer = MousePointerConstants.vbSizeAll
            End If
        Else
            parent.picPreview.MousePointer = MousePointerConstants.vbDefault
        End If
        
        parent.picPreview.Refresh
    End If
    
        
End Sub

Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    Dim curX As Integer, curY As Integer
    Dim zoom As Single
    zoom = parent.magnifier.zoom
        
    curX = (parent.Hscr.value + X) / zoom
    curY = (parent.Vscr.value + Y) / zoom
    
    If State = S_SelectingSingle Then
        Dim lvzidx As Integer, objidx As Long
        If getMapObjectAtPos(curX, curY, lvzidx, objidx) <> -1 Then
            If IsControl(Shift) Then
                If lvz(lvzidx).mapobjects(objidx).selected Then
                    'Remove it from selection
                    Call RemoveFromSelection(lvzidx, objidx)
                Else
                    'Add it to selection
                    Call AddToSelection(lvzidx, objidx)
                End If
            Else
                Call EmptySelection
                Call AddToSelection(lvzidx, objidx)
            End If
        End If
    ElseIf State = S_SelectingArea Then
        
        'Get all objects in that area
        Dim hoverobjs() As MapObjectRef
        Dim count As Long
        count = getMapObjectsInArea(intMinimum(firstX, curX), intMinimum(firstY, curY), intMaximum(firstX, curX), intMaximum(firstY, curY), curX < firstX, hoverobjs)
        
        'Select them
        For i = 0 To count - 1
            Call AddToSelection(hoverobjs(i).lvzidx, hoverobjs(i).objidx)
        Next
    End If
    
    
    parent.picPreview.MousePointer = MousePointerConstants.vbDefault
    
    Call parent.UpdatePreview
    
    State = S_Nothing
End Sub




Private Function CountTotalImageDefinitions() As Integer
    Dim i As Integer
    Dim total As Integer
    total = 0
    For i = 0 To lvzcount - 1
        total = total + lvz(i).ImageDefinitionCount
    Next
    CountTotalImageDefinitions = total
End Function



Property Get MapObjectDefaultLayer() As LVZLayerEnum
    MapObjectDefaultLayer = defaultLayer
End Property

Property Let MapObjectDefaultLayer(n_layer As LVZLayerEnum)
    defaultLayer = n_layer
End Property

Property Get MapObjectDefaultMode() As LVZModeEnum
    MapObjectDefaultMode = defaultDisplay
End Property

Property Let MapObjectDefaultMode(n_mode As LVZModeEnum)
    defaultDisplay = n_mode
End Property








Function getImageDefFromCoordinates(ByVal X As Integer, ByVal Y As Integer, ByRef lvzidx As Integer, ByRef imgidx As Integer, totalWidth As Integer, totalHeight As Integer, scrollValue As Integer) As Long

    Dim xOrg As Integer, yOrg As Integer

    
    Dim i As Integer, j As Integer, c As Integer
    c = 0
    yOrg = (c \ (totalWidth \ ImageTilesetW)) * ImageTilesetH - scrollValue * ImageTilesetW
    
    For i = 0 To lvzcount - 1
        For j = 0 To lvz(i).ImageDefinitionCount - 1
            
            If X >= xOrg And X <= xOrg + ImageTilesetW And _
               Y >= yOrg And Y <= yOrg + ImageTilesetH Then
               
                lvzidx = i
                imgidx = j
               
                getImageDefFromCoordinates = i * 65536 + j
'                frmGeneral.Label6.Caption = c & " " & i & " " & j
                Exit Function
            End If
            
            
            c = c + 1
            xOrg = (xOrg + ImageTilesetW) Mod totalWidth
            yOrg = (c \ (totalWidth \ ImageTilesetW)) * ImageTilesetH - scrollValue * ImageTilesetW
        Next
    Next

    getImageDefFromCoordinates = -1
    
End Function

Function HaveLVZ() As Boolean
    HaveLVZ = (lvzcount > 0)
End Function

Function hasImageDefinitions() As Boolean
    Dim i As Integer
    For i = 0 To lvzcount - 1
        If lvz(i).ImageDefinitionCount > 0 Then
            hasImageDefinitions = True
            Exit Function
        End If
    Next
    
    hasImageDefinitions = False
End Function

Function isImageDependedOn(lvzidx As Integer, imgidx As Integer) As Boolean
    Dim i As Long
    For i = 0 To lvz(lvzidx).mapObjectCount - 1
        If lvz(lvzidx).mapobjects(i).imgidx = imgidx Then
            isImageDependedOn = True
            Exit Function
        End If
    Next
    
    For i = 0 To lvz(lvzidx).ScreenobjectCount - 1
        If lvz(lvzidx).screenobjects(i).imgidx = imgidx Then
            isImageDependedOn = True
            Exit Function
        End If
    Next

    isImageDependedOn = False
End Function

Function isFileDependedOn(filename As String) As Boolean
    Dim i As Integer, j As Integer
    
    For j = 0 To lvzcount - 1
        For i = 0 To lvz(j).ImageDefinitionCount - 1
            If lvz(j).imagedefinitions(i).imagename = filename Then
                isFileDependedOn = True
                Exit Function
            End If
        Next
    Next

    isFileDependedOn = False
End Function

Function CountImagesUsingFile(filename As String) As Long
    Dim i As Integer, j As Integer
    Dim count As Long
    
    count = 0
    
    For j = 0 To lvzcount - 1
        For i = 0 To lvz(j).ImageDefinitionCount - 1
            If lvz(j).imagedefinitions(i).imagename = filename Then
                count = count + 1
            End If
        Next
    Next

    CountImagesUsingFile = count
End Function

Private Sub ReloadImagesUsingFile(filename As String)
    Dim i As Integer, j As Integer
    
    For j = 0 To lvzcount - 1
        For i = 0 To lvz(j).ImageDefinitionCount - 1
            
            If lvz(j).imagedefinitions(i).imagename = filename Then
                
                Call buildSingleLVZImage(lvz(j).imagedefinitions(i).picboxIdx, j, i)
                
'                With lvz(j).imagedefinitions(i)
'
'
'                    parent.picLVZItem(.picboxidx).AutoSize = True
'
'                    Call LoadPic(parent.picLVZItem(.picboxidx), SearchFile(.imagename))
'
'                    Call buildZoomLVZImages(.picboxidx)
'
'                    parent.picLVZItem(.picboxidx).AutoSize = False
'
'                    .picwidth = parent.picLVZItem(.picboxidx).width \ .animationFramesX
'                    .picheight = parent.picLVZItem(.picboxidx).height \ .animationFramesY
'
'                    Call buildLibraryImages(j, i)
'
'                End With
            End If
        Next
    Next

End Sub


Function CountMapObjectsUsingImage(lvzidx As Integer, imgidx As Integer) As Long
    Dim i As Long
    Dim count As Long
    
    count = 0
    
    With lvz(lvzidx)
    
        For i = 0 To .mapObjectCount - 1
            If .mapobjects(i).imgidx = imgidx Then
                count = count + 1
            End If
        Next

    End With
    
    CountMapObjectsUsingImage = count
End Function

Function CountScreenObjectsUsingImage(lvzidx As Integer, imgidx As Integer) As Long
    Dim i As Long
    Dim count As Long
    
    count = 0
    
    With lvz(lvzidx)
    
        For i = 0 To .ScreenobjectCount - 1
            If .screenobjects(i).imgidx = imgidx Then
                count = count + 1
            End If
        Next
    
    End With
    
    CountScreenObjectsUsingImage = count
End Function


 Sub RemoveLinksToImage(lvzidx As Integer, imgidx As Integer)
    Dim i As Long
    
    
    'We have to use a 'While' loop, because everytime a map object is removed, the count goes down 1
    'With a For loop, the 'i' would increment everytime and skip objects that are just after those removed
    
    With lvz(lvzidx)
    
        i = 0
        
        While i < .mapObjectCount
            If .mapobjects(i).imgidx = imgidx Then
                Call removeMapObjectFromLVZ(lvzidx, i)
            Else
                i = i + 1
            End If
        Wend
        
        i = 0
        
        While i < .ScreenobjectCount
            If .screenobjects(i).imgidx = imgidx Then
                Call removeScreenObjectFromLVZ(lvzidx, i)
            Else
                i = i + 1
            End If
        Wend
    
    
    End With
    


End Sub

Sub RemoveLinksToFile(filename As String)
    Dim i As Integer, j As Integer
    
    
    For j = 0 To lvzcount - 1
        For i = 0 To lvz(j).ImageDefinitionCount - 1
            If lvz(j).imagedefinitions(i).imagename = filename Then
                
                Call RemoveLinksToImage(j, i)
'                If deleteImage Then
                Call removeImageDefinitionFromLVZ(j, i)
'                Else
'                    Call lvz.setImageDefinitionFile(j, i, "")
'                End If
            End If
        Next
    Next
    
End Sub


Sub MoveSelection(dx As Integer, dy As Integer)
    Dim i As Long
    
    For i = 0 To selectedObjectsCount - 1
        With selectedObjects(i)
            With lvz(.lvzidx).mapobjects(.objidx)
            
                .X = .X + dx
                .Y = .Y + dy
                
            End With
        End With
    
    Next
    
    dontCallSelectionChange = True
    If selCoordX <> MULTIPLE_VALUES Then
        selCoordX = selCoordX + dx
        frmGeneral.txtLvzX.Text = selCoordX
    End If
    
    If selCoordY <> MULTIPLE_VALUES Then
        selCoordY = selCoordY + dy
        frmGeneral.txtLvzY.Text = selCoordY
    End If
    dontCallSelectionChange = False
    
'    Call ShowSelection(True)

End Sub



Private Sub HighLightObjects(ByRef mapobjs() As MapObjectRef, count As Long, color As Long)
    Dim i As Long
    Dim zoom As Single
    Dim hDC As Long
    
    Dim offsX As Integer, offsY As Integer, picWidth As Integer, picHeight As Integer
    
    Dim lvzidx As Integer, objidx As Long
    
    With parent
        zoom = .magnifier.zoom
        hDC = .picPreview.hDC
        
        offsX = .Hscr.value / zoom
        offsY = .Vscr.value / zoom
    End With
    
    For i = 0 To count - 1
    
        With mapobjs(i)
            lvzidx = .lvzidx
            objidx = .objidx
        End With
            
        With lvz(lvzidx)
            With .imagedefinitions(.mapobjects(objidx).imgidx)
                picWidth = .picWidth
                picHeight = .picHeight
            End With
            
            With .mapobjects(objidx)

                DrawRectangle hDC, Int((.X - offsX) * zoom), Int((.Y - offsY) * zoom), Int((.X - offsX + picWidth) * zoom) - 1, Int((.Y - offsY + picHeight) * zoom) - 1, color
                
            End With
        End With
            
    
    Next
End Sub

Sub ShowSelection(updateText As Boolean)
    

    dontCallSelectionChange = True
    
    If selectedObjectsCount > 0 Then
        
        Call HighLightObjects(selectedObjects, selectedObjectsCount, vbMagenta)
        
        If updateText Then
            With frmGeneral
                .txtLVZDisplayTime.Enabled = True
                .txtLVZDisplayTime.Text = selDisplayTime

                .cmbLvzDisplayType.Enabled = True
                .cmbLvzDisplayType.ListIndex = selDisplayMode

                .cmbLvzLayerType.Enabled = True
                .cmbLvzLayerType.ListIndex = selLayer

                .txtLvzObjectID.Enabled = True
                .txtLvzObjectID.Text = selObjectID
              
              If selCoordX = MULTIPLE_VALUES Then
                .txtLvzX.Enabled = False
                .txtLvzX.Text = "0"
              Else
                .txtLvzX.Enabled = True
                .txtLvzX.Text = selCoordX
              End If
              
              If selCoordY = MULTIPLE_VALUES Then
                .txtLvzY.Enabled = False
                .txtLvzY.Text = "0"
              Else
                .txtLvzY.Enabled = True
                .txtLvzY.Text = selCoordY
              End If
            End With
        End If
                    
    ElseIf updateText Then
        'No selection
        With frmGeneral
            .txtLVZDisplayTime.Enabled = False
            .txtLVZDisplayTime.Text = ""

            .cmbLvzDisplayType.Enabled = False
            .cmbLvzDisplayType.ListIndex = -1

            .cmbLvzLayerType.Enabled = False
            .cmbLvzLayerType.ListIndex = -1

            .txtLvzObjectID.Enabled = False
            .txtLvzObjectID.Text = ""

          .txtLvzX.Enabled = False
          .txtLvzX.Text = ""
          .txtLvzY.Enabled = False
          .txtLvzY.Text = ""
        End With
    End If
    dontCallSelectionChange = False

End Sub




Function ChangeMapObjectLayer(ByVal lvzidx As Integer, ByVal objidx As Long, ByVal newlayer As LVZLayerEnum) As Long
    'Moves an object of layer, and keeps it sorted
    'Returns the new index of the object
    
    Dim tmpobj As LVZMapObject
    Dim newidx As Long
    
    With lvz(lvzidx)

        
        'We have to move it in the array
        tmpobj = .mapobjects(objidx)
        
        If tmpobj.layer = newlayer Then
            'No layer change
            ChangeMapObjectLayer = objidx
            Exit Function
        End If
        
        'Make sure the new layer is set before inserting, so we insert it at the correct index
        tmpobj.layer = newlayer
        
        Dim wasSelected As Boolean
'        wasSelected = tmpobj.selected
'        tmpobj.selected = False
        
        newidx = AddMapObjectToLVZ(lvzidx, tmpobj)
'        Call InsertMapObjectAtIdx(lvzidx, tmpobj, .nextIndexFor(newlayer))
        
'        .mapobjects(newidx).selected = wasSelected
        
        
        'Make sure the layer of the OLD object is still not changed,
        'else the 'nextIndexFor' won't be updated correctly
        
        
        If newidx <= objidx Then objidx = objidx + 1 Else newidx = newidx - 1
        
        Call removeMapObjectFromLVZ(lvzidx, objidx)
        

        
        ChangeMapObjectLayer = newidx
        
        
    End With
    
End Function






Sub ChangeSelectionLayer(newlayer As LVZLayerEnum)
    If HasSelection And Not dontCallSelectionChange Then
'        Dim oldlayer As LVZLayerEnum

        
        
        selLayer = newlayer
        
        'THIS DOESN'T WORK
        'because once the first object is changed layer, it moves the other objects in the array
        'so the references in selectedObjects() become obsolete
        Dim i As Long

        For i = 0 To selectedObjectsCount - 1

            With selectedObjects(i)
                .objidx = ChangeMapObjectLayer(.lvzidx, .objidx, newlayer)
            End With
        Next
        
        
        
        
        
        
        Call RedrawLVZLayers
        

    End If
End Sub

Sub ChangeSelectionDisplayTime(newtime As Long)
    If HasSelection Then
        Dim i As Long
        
        selDisplayTime = newtime
        
        For i = 0 To selectedObjectsCount - 1
            With selectedObjects(i)
                lvz(.lvzidx).mapobjects(.objidx).displayTime = newtime
            End With
        Next
        
        Call ShowSelection(True)
        
    End If
End Sub

Sub ChangeSelectionDisplayMode(newmode As LVZModeEnum)
    If HasSelection Then
        Dim i As Long
        
        selDisplayMode = newmode
        
        For i = 0 To selectedObjectsCount - 1
            With selectedObjects(i)
                lvz(.lvzidx).mapobjects(.objidx).mode = newmode
            End With
        Next
        
        Call ShowSelection(True)
        
    End If
End Sub

Sub ChangeSelectionObjectID(newID As Integer)
    If HasSelection Then
        Dim i As Long
        
        selObjectID = newID
        
        For i = 0 To selectedObjectsCount - 1
            With selectedObjects(i)
                lvz(.lvzidx).mapobjects(.objidx).objectID = newID
            End With
        Next
        
        Call ShowSelection(True)
    End If
End Sub

Sub ChangeSelectionX(newX As Integer)
    If HasSelection And newX <> MULTIPLE_VALUES Then
        Dim i As Long
        For i = 0 To selectedObjectsCount - 1
            With selectedObjects(i)
                lvz(.lvzidx).mapobjects(.objidx).X = newX
            End With
        Next
        
        selCoordX = newX
        
        Call RedrawLVZLayers
    End If
End Sub

Sub ChangeSelectionY(newY As Integer)
    If HasSelection And newY <> MULTIPLE_VALUES Then
        Dim i As Long
        For i = 0 To selectedObjectsCount - 1
            With selectedObjects(i)
                lvz(.lvzidx).mapobjects(.objidx).Y = newY
            End With
        Next
        
        selCoordY = newY
        Call RedrawLVZLayers
    End If
End Sub


Sub DeleteSelectedObject()
    If HasSelection Then
'        Call removeMapObjectFromLVZ(SelLVZIdx, selObjIdx)
'
        Dim i As Long
        
        For i = 0 To selectedObjectsCount - 1
            With selectedObjects(i)
                Call removeMapObjectFromLVZ(.lvzidx, .objidx)
            End With
        Next

        Call EmptySelection
        
        Call RedrawLVZLayers
    End If
End Sub

Property Get snap() As Integer
    snap = c_Snap
End Property

Property Let snap(newsnap As Integer)
    c_Snap = newsnap
    If c_Snap <= 0 Then c_Snap = 1
End Property

Private Sub Class_Terminate()
    Set parent = Nothing
    Set undoch = Nothing
    
    Dim i As Integer
    
    For i = 0 To c_maxidx
'        Set picLVZImagesZoomPxl(i) = Nothing
'        Set picLVZImagesZoom1_4(i) = Nothing
'        Set picLVZImagesZoom1_2(i) = Nothing
        Set picLVZImages(i) = Nothing
        Set picLVZImagesZoom(i) = Nothing
'        Set picLVZImagesZoom2_1(i) = Nothing
        Set picLVZImagesLibrary(i) = Nothing
    Next
End Sub




Friend Sub DrawLVZsInArea(hDC As Long, layer As Integer, area As RECT, firstscreenX As Integer, firstscreenY As Integer)
    'left, top, right, bottom: World coordinates (pixels)
    'firstscreenX, firstscreenY: Screen coordinates corresponding to (left, top)
    
    Dim lvzidx As Integer
    Dim objidx As Long, _
        minobj As Long, maxobj As Long
    
    Dim objarea As RECT, intersect As RECT
    
    Dim srchDC As Long, _
        DestX As Integer, DestY As Integer, _
        width As Integer, height As Integer, _
        SrcX As Integer, SrcY As Integer, _
        picboxID As Integer
    
    Dim zoom As Single
    zoom = parent.magnifier.zoom
    
    Dim UsingPixels As Boolean
    UsingPixels = parent.magnifier.UsingPixels
    
    Dim objW As Integer, objH As Integer
    
                
    For lvzidx = 0 To lvzcount - 1
        With lvz(lvzidx)
            'Index of the first object on this layer
            If layer > lyr_BelowAll Then minobj = .nextIndexFor(layer - 1) Else minobj = 0

            'Index of the last object on this layer
            maxobj = .nextIndexFor(layer) - 1
            
            For objidx = minobj To maxobj
                With .mapobjects(objidx)
                    
                    objW = GetMapObjectWidth(lvzidx, objidx)
                    objH = GetMapObjectHeight(lvzidx, objidx)
                    
                    objarea.Left = .X
                    objarea.Right = .X + objW - 1
                    objarea.Top = .Y
                    objarea.Bottom = .Y + objH - 1
                                            
                    If Intersection(intersect, area, objarea) Then
                        'Get the source hDC we need
                        picboxID = lvz(lvzidx).imagedefinitions(.imgidx).picboxIdx
                        
                        srchDC = pichDCzoom(picboxID, zoom)

                        
                        'Width / height to draw
                        width = (intersect.Right - intersect.Left + 1) * zoom
                        height = (intersect.Bottom - intersect.Top + 1) * zoom
                        DestX = firstscreenX + (intersect.Left - area.Left) * zoom
                        DestY = firstscreenY + (intersect.Top - area.Top) * zoom
                        
                        If width > 0 And height > 0 Then
                            If DestX = firstscreenX Then SrcX = (area.Left - objarea.Left) * zoom Else SrcX = 0
                            If DestY = firstscreenY Then SrcY = (area.Top - objarea.Top) * zoom Else SrcY = 0
                            
                            
                            'If we're too far zoomed out, it's possible we don't have anything to draw for small images
                            TransparentBlt hDC, DestX, DestY, width, height, srchDC, SrcX, SrcY, vbBlack
                            
'                            If .selected Then
'                                DrawRectangle hDC, firstscreenX + (objarea.Left - area.Left) * zoom, firstscreenY + (objarea.Top - area.Top) * zoom, firstscreenX + (objarea.Right + 1 - area.Left) * zoom, firstscreenY + (objarea.Bottom + 1 - area.Top) * zoom, frmgeneral.TilesetBackgroundColor
''                                DrawRectangleHdc hdc, DestX, DestY, DestX + width, DestY + height
'                            End If
                        End If
                    End If
                End With
            Next
        End With
        
    Next
End Sub












Sub DoEditImageDefinitionPic(lvzidx As Integer, imgidx As Integer)
        
    Dim picboxIdx As Integer
    picboxIdx = lvz(lvzidx).imagedefinitions(imgidx).picboxIdx
        
    Dim targetDC As Long
    targetDC = picLVZImages(picboxIdx).hDC
    
        
'     If EditImage(frmGeneral, myPic.hDC, 0, 0, myPic.width, myPic.height, False, myPic) Then
    If EditImage(frmGeneral, targetDC, 0, 0, picLVZImages(picboxIdx).width, picLVZImages(picboxIdx).height, False) Then
     
'               Dim findex As Long
'
'               Call buildZoomLVZImages(myPic.Index)
'
'               With lvz(lvzidx).imagedefinitions(imgidx)
'                   .picwidth = myPic.width \ .animationFramesX
'                   .picheight = myPic.height \ .animationFramesY
'
'                   findex = getIndexOfFile(lvzidx, .imagename)
'
'                   If CountImagesUsingFile(.imagename) > 1 Then
'                       'More than one image defs using this file
'
'                       If MessageBox("More than one image definition is using the picture you edited. Do you want to apply changes to all these image definitions? If you answer No, a new file will be created for this image definition.", vbQuestion + vbYesNo) = vbNo Then
'
'                           Dim newpath As String
'                           newpath = GetUniqueFilename(GetPathTo(lvz(lvzidx).files(findex).path), .imagename)
'
'                           findex = AddFileToLVZ(lvzidx, newpath)
'
'                           .imagename = GetFileTitle(newpath)
'
'
'
'                       End If
'                   End If
'               End With
'
'
'
'
'               If findex <> -1 Then
'                   myPic.Picture = myPic.Image
'
'                   Call SavePicture(myPic.Picture, lvz(lvzidx).files(findex).path)
'
'                   Call ReloadImagesUsingFile(lvz(lvzidx).imagedefinitions(imgidx).imagename)
'               Else
'                   MessageBox "Error updating file!", vbExclamation
'               End If
'
'              Call parent.tileset.SelectLVZ(vbLeftButton, lvzidx, imgidx, True)
     End If
    

    Call parent.UpdateLevel
End Sub

Sub DoEditImageDefinitionProperties(lvzidx As Integer, imgidx As Integer)
    
    With lvz(lvzidx).imagedefinitions(imgidx)
    
         .animationTime = InputBoxValue("Animation time (in 1/100th of second) = ", "Edit Image Definition", .animationTime, 1, 32767)
    
        .animationFramesX = InputBoxValue("Animation width (number of frames) = ", "Edit Image Definition", .animationFramesX, 1, 32767)
        
        .animationFramesY = InputBoxValue("Animation height (number of frames) = ", "Edit Image Definition", .animationFramesY, 1, 32767)
        
        Call parent.tileset.SelectLVZ(vbLeftButton, lvzidx, imgidx, True)
        
        Call parent.UpdateLevel
    End With
End Sub




Function GetGlobalIndexOfImageDefinition(lvzidx As Integer, imgidx As Integer) As Integer
    'Returns the index from 0 to 255 of an image definition
    Dim gIdx As Integer
    Dim i As Integer, j As Integer
    
    gIdx = -1
    
    For i = 0 To lvzidx - 1
        gIdx = gIdx + lvz(i).ImageDefinitionCount
    Next
    
    If lvz(lvzidx).ImageDefinitionCount > imgidx Then
        gIdx = gIdx + imgidx
    Else
        gIdx = gIdx - 1
    End If
    
    GetGlobalIndexOfImageDefinition = gIdx
End Function

Sub GetLocalIndexOfImageDefinition(ByVal globalIndex As Integer, ByRef lvzidx As Integer, ByRef imgidx As Integer)
    'Returns lvzidx and imgidx of the given image definition
    'Returns the index from 0 to 255 of an image definition
    Dim cumul As Integer, i As Integer, rest As Integer
    
    i = 0
    cumul = -1
    
    While cumul < globalIndex
        
        If globalIndex - cumul >= lvz(i).ImageDefinitionCount Then
            cumul = cumul + lvz(i).ImageDefinitionCount
        Else
        
            rest = globalIndex - cumul
            cumul = globalIndex
            
        End If
        
        i = i + 1
        
    Wend
        
    lvzidx = i
    imgidx = rest
End Sub

Public Property Get GetTotalImageDefinitions() As Integer
    GetTotalImageDefinitions = totalImageDefinitions
End Property
