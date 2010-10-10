VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "boolArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim columnCount As Long

Dim rowCount As Long

Dim totalBytes As Long

Dim field() As Byte
Dim field_ptr As Long

    ''''''''''''''''''''''
    '   32   2  16
    '    4  [X]  1
    '   64   8  128
Public Enum Neighbor
    N_TopLeft = 32
    N_Top = 2
    N_TopRight = 16
    N_Left = 4
    N_Right = 1
    N_BottomLeft = 64
    N_Bottom = 8
    N_BottomRight = 128
End Enum
    
Sub InitArray(n_columnsCount As Long, n_rowsCount As Long, default As Byte)
    Dim i As Long, j As Long
    
    columnCount = n_columnsCount
    rowCount = n_rowsCount
       
    ReDim field(columnCount * rowCount - 1)
    
    field_ptr = VarPtr(field(0))
    totalBytes = columnCount * rowCount
'
'    messagebox field_ptr & " " & VarPtr(field(1, 0)) & vbCrLf & VarPtr(field(0, 1))
'    messagebox 0 & " " & VarPtr(field(1, 0)) - field_ptr & vbCrLf & VarPtr(field(0, 1)) - field_ptr
'
    If default Then FillMemory ByVal field_ptr, totalBytes, default
        
End Sub

Sub FillArray(value As Byte)
    FillMemory ByVal field_ptr, totalBytes, value
End Sub


Property Get value(col As Integer, row As Integer) As Boolean
    value = field(col + row * columnCount)
End Property

Property Let value(col As Integer, row As Integer, val As Boolean)
    field(col + row * columnCount) = IIf(val, 1, 0)
End Property

Property Get valueL(col As Long, row As Long) As Boolean
    valueL = field(col + row * columnCount)
End Property

Property Let valueL(col As Long, row As Long, val As Boolean)
    field(col + row * columnCount) = IIf(val, 1, 0)
End Property

Property Get valueB(col As Long, row As Long) As Byte
    valueB = field(col + row * columnCount)
End Property

Property Let valueB(col As Long, row As Long, val As Byte)
    field(col + row * columnCount) = val
End Property




Public Property Get field_pointer() As Long
    field_pointer = field_ptr
End Property

Public Property Get field_length() As Long
    field_length = totalBytes
End Property

Public Property Get width() As Long
    width = columnCount
End Property

Public Property Get height() As Long
    height = rowCount
End Property

Sub MoveRectangle(firstcol As Integer, lastcol As Integer, firstrow As Integer, lastrow As Integer, dx As Integer, dy As Integer, eraseSource As Boolean)
    'firstcol, lastcol, firstrow, lastrow: Area to move/copy
    'dX, dY: where to copy/move the area
    'eraseSource: set the source to 0?
    
    Dim r As Long
    Dim src_ptr As Long, dest_ptr As Long
    Dim width As Long, height As Long
    Dim Buffer() As Byte
    
    src_ptr = field_ptr + (firstrow * columnCount) + firstcol
    dest_ptr = src_ptr + dx + (dy * columnCount)
    
    width = lastcol - firstcol + 1
    height = lastrow - firstrow + 1
    
    Dim useBuffer As Boolean
    Dim sourceWidth As Long
    sourceWidth = columnCount
    
    'Don't bother using a buffer if the destination does not overlap the source
    useBuffer = (Abs(dx) < width And Abs(dy) < height)
    
    
    If useBuffer Then
        'Fill buffer
        Dim buf_ptr As Long
        
        ReDim Buffer(width * height - 1)
        buf_ptr = VarPtr(Buffer(0))
        
        For r = firstrow To lastrow
            CopyMemory ByVal buf_ptr, ByVal src_ptr, width
            
            'Clear source if needed
            If eraseSource Then ZeroMemory ByVal src_ptr, width
            
            buf_ptr = buf_ptr + width 'Increment to the next row
            src_ptr = src_ptr + columnCount 'Increment to the next row
        Next r
        
        'reset the source pointer to point to the buffer
        src_ptr = VarPtr(Buffer(0))
        sourceWidth = width
    End If
    
    'Copy values
    
    'Don't erase source if we were using a buffer; it's already deleted
    'This also avoid making the boolean check on every row
    eraseSource = eraseSource And (Not useBuffer)
    
    For r = firstrow To lastrow
        CopyMemory ByVal dest_ptr, ByVal src_ptr, width
        
        If eraseSource Then ZeroMemory ByVal src_ptr, width
        
        dest_ptr = dest_ptr + columnCount 'Increment to the next row
        src_ptr = src_ptr + sourceWidth   'Increment to the next row
    Next r


End Sub

Sub SetRectangle(firstcol As Integer, lastcol As Integer, firstrow As Integer, lastrow As Integer, val As Byte)
    Dim r As Long
    Dim cur_ptr As Long
    Dim Length As Long
    
    If firstcol < 0 Or lastcol >= columnCount Or firstrow < 0 Or lastrow >= rowCount Then
        AddDebug "ERROR in boolArray.SetRectangle: " & firstcol & "," & firstrow & " - " & lastcol & "," & lascol
        Exit Sub
    End If
    
    cur_ptr = field_ptr + (firstrow * columnCount) + firstcol
    Length = lastcol - firstcol + 1
    
    If val Then
        For r = firstrow To lastrow
            FillMemory ByVal cur_ptr, Length, val
            cur_ptr = cur_ptr + columnCount
        Next r
    Else
        For r = firstrow To lastrow
            ZeroMemory ByVal cur_ptr, Length
            cur_ptr = cur_ptr + columnCount
        Next r
    End If
End Sub



Sub SetFromBooleanArray(ByRef bools() As Boolean, width As Long, height As Long)
    Dim i As Long, j As Long
    
    Call InitArray(width, height, 0)
    
    For i = 0 To width - 1
        For j = 0 To height - 1
            If bools(i, j) Then
                field(i + j * columnCount) = 1
            Else
                field(i + j * columnCount) = 0
            End If
        Next
    Next
End Sub

Sub CopyFromBitfield(ByRef srcfield As boolArray)
    If srcfield.field_length <> totalBytes Then
        'Incompatible fields
        Call InitArray(srcfield.width, srcfield.height, 0)
    End If
    
    CopyMemory ByVal field_ptr, ByVal srcfield.field_pointer, totalBytes
End Sub

Function GetNeighborhood(ByVal i As Integer, ByVal j As Integer, diagonals As Boolean) As Byte
    ''''''''''''''''''''''
    '    1   2   4
    '    8  [X] 16
    '   32  64 128
    Dim chkLeft As Boolean, chkTop As Boolean, chkRight As Boolean, chkBottom As Boolean
    chkRight = i < columnCount - 1
    chkLeft = i > 0
    chkBottom = j < rowCount - 1
    chkTop = j > 0
    
    If chkLeft Then GetNeighborhood = GetNeighborhood + (Not value(i - 1, j) And Neighbor.N_Left)
    If chkTop Then GetNeighborhood = GetNeighborhood + (Not value(i, j - 1) And Neighbor.N_Top)
    If chkRight Then GetNeighborhood = GetNeighborhood + (Not value(i + 1, j) And Neighbor.N_Right)
    If chkBottom Then GetNeighborhood = GetNeighborhood + (Not value(i, j + 1) And Neighbor.N_Bottom)
    
    'Diagonals
    If diagonals Then
        If chkLeft And chkTop Then GetNeighborhood = GetNeighborhood + (value(i - 1, j - 1) And Neighbor.N_TopLeft)
        If chkLeft And chkBottom Then GetNeighborhood = GetNeighborhood + (value(i - 1, j + 1) And Neighbor.N_BottomLeft)
        If chkRight And chkTop Then GetNeighborhood = GetNeighborhood + (value(i + 1, j - 1) And Neighbor.N_TopRight)
        If chkRight And chkBottom Then GetNeighborhood = GetNeighborhood + (value(i + 1, j + 1) And Neighbor.N_BottomRight)
    End If
        
End Function

Private Sub Class_Terminate()
    'Erase field
    ReDim field(0)
    
End Sub
