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
          
10        columnCount = n_columnsCount
20        rowCount = n_rowsCount
             
30        ReDim field(columnCount * rowCount - 1)
          
40        field_ptr = VarPtr(field(0))
50        totalBytes = columnCount * rowCount
      '
      '    messagebox field_ptr & " " & VarPtr(field(1, 0)) & vbCrLf & VarPtr(field(0, 1))
      '    messagebox 0 & " " & VarPtr(field(1, 0)) - field_ptr & vbCrLf & VarPtr(field(0, 1)) - field_ptr
      '
60        If default Then FillMemory ByVal field_ptr, totalBytes, default
              
End Sub

Sub FillArray(value As Byte)
10        FillMemory ByVal field_ptr, totalBytes, value
End Sub


Property Get value(col As Integer, row As Integer) As Boolean
10        value = field(col + row * columnCount)
End Property

Property Let value(col As Integer, row As Integer, val As Boolean)
10        field(col + row * columnCount) = IIf(val, 1, 0)
End Property

Property Get valueL(col As Long, row As Long) As Boolean
10        valueL = field(col + row * columnCount)
End Property

Property Let valueL(col As Long, row As Long, val As Boolean)
10        field(col + row * columnCount) = IIf(val, 1, 0)
End Property

Property Get valueB(col As Long, row As Long) As Byte
10        valueB = field(col + row * columnCount)
End Property

Property Let valueB(col As Long, row As Long, val As Byte)
10        field(col + row * columnCount) = val
End Property




Public Property Get field_pointer() As Long
10        field_pointer = field_ptr
End Property

Public Property Get field_length() As Long
10        field_length = totalBytes
End Property

Public Property Get width() As Long
10        width = columnCount
End Property

Public Property Get height() As Long
10        height = rowCount
End Property

Sub MoveRectangle(firstcol As Integer, lastcol As Integer, firstrow As Integer, lastrow As Integer, dx As Integer, dy As Integer, eraseSource As Boolean)
          'firstcol, lastcol, firstrow, lastrow: Area to move/copy
          'dX, dY: where to copy/move the area
          'eraseSource: set the source to 0?
          
          Dim r As Long
          Dim src_ptr As Long, dest_ptr As Long
          Dim width As Long, height As Long
          Dim Buffer() As Byte
          
10        src_ptr = field_ptr + (firstrow * columnCount) + firstcol
20        dest_ptr = src_ptr + dx + (dy * columnCount)
          
30        width = lastcol - firstcol + 1
40        height = lastrow - firstrow + 1
          
          Dim useBuffer As Boolean
          Dim sourceWidth As Long
50        sourceWidth = columnCount
          
          'Don't bother using a buffer if the destination does not overlap the source
60        useBuffer = (Abs(dx) < width And Abs(dy) < height)
          
          
70        If useBuffer Then
              'Fill buffer
              Dim buf_ptr As Long
              
80            ReDim Buffer(width * height - 1)
90            buf_ptr = VarPtr(Buffer(0))
              
100           For r = firstrow To lastrow
110               CopyMemory ByVal buf_ptr, ByVal src_ptr, width
                  
                  'Clear source if needed
120               If eraseSource Then ZeroMemory ByVal src_ptr, width
                  
130               buf_ptr = buf_ptr + width 'Increment to the next row
140               src_ptr = src_ptr + columnCount 'Increment to the next row
150           Next r
              
              'reset the source pointer to point to the buffer
160           src_ptr = VarPtr(Buffer(0))
170           sourceWidth = width
180       End If
          
          'Copy values
          
          'Don't erase source if we were using a buffer; it's already deleted
          'This also avoid making the boolean check on every row
190       eraseSource = eraseSource And (Not useBuffer)
          
200       For r = firstrow To lastrow
210           CopyMemory ByVal dest_ptr, ByVal src_ptr, width
              
220           If eraseSource Then ZeroMemory ByVal src_ptr, width
              
230           dest_ptr = dest_ptr + columnCount 'Increment to the next row
240           src_ptr = src_ptr + sourceWidth   'Increment to the next row
250       Next r


End Sub

Sub SetRectangle(firstcol As Integer, lastcol As Integer, firstrow As Integer, lastrow As Integer, val As Byte)
          Dim r As Long
          Dim cur_ptr As Long
          Dim Length As Long
          
10        If firstcol < 0 Or lastcol >= columnCount Or firstrow < 0 Or lastrow >= rowCount Then
20            AddDebug "ERROR in boolArray.SetRectangle: " & firstcol & "," & firstrow & " - " & lastcol & "," & lascol
30            Exit Sub
40        End If
          
50        cur_ptr = field_ptr + (firstrow * columnCount) + firstcol
60        Length = lastcol - firstcol + 1
          
70        If val Then
80            For r = firstrow To lastrow
90                FillMemory ByVal cur_ptr, Length, val
100               cur_ptr = cur_ptr + columnCount
110           Next r
120       Else
130           For r = firstrow To lastrow
140               ZeroMemory ByVal cur_ptr, Length
150               cur_ptr = cur_ptr + columnCount
160           Next r
170       End If
End Sub



Sub SetFromBooleanArray(ByRef bools() As Boolean, width As Long, height As Long)
          Dim i As Long, j As Long
          
10        Call InitArray(width, height, 0)
          
20        For i = 0 To width - 1
30            For j = 0 To height - 1
40                If bools(i, j) Then
50                    field(i + j * columnCount) = 1
60                Else
70                    field(i + j * columnCount) = 0
80                End If
90            Next
100       Next
End Sub

Sub CopyFromBitfield(ByRef srcfield As boolArray)
10        If srcfield.field_length <> totalBytes Then
              'Incompatible fields
20            Call InitArray(srcfield.width, srcfield.height, 0)
30        End If
          
40        CopyMemory ByVal field_ptr, ByVal srcfield.field_pointer, totalBytes
End Sub

Function GetNeighborhood(ByVal i As Integer, ByVal j As Integer, diagonals As Boolean) As Byte
          ''''''''''''''''''''''
          '    1   2   4
          '    8  [X] 16
          '   32  64 128
          Dim chkLeft As Boolean, chkTop As Boolean, chkRight As Boolean, chkBottom As Boolean
10        chkRight = i < columnCount - 1
20        chkLeft = i > 0
30        chkBottom = j < rowCount - 1
40        chkTop = j > 0
          
50        If chkLeft Then GetNeighborhood = GetNeighborhood + (Not value(i - 1, j) And Neighbor.N_Left)
60        If chkTop Then GetNeighborhood = GetNeighborhood + (Not value(i, j - 1) And Neighbor.N_Top)
70        If chkRight Then GetNeighborhood = GetNeighborhood + (Not value(i + 1, j) And Neighbor.N_Right)
80        If chkBottom Then GetNeighborhood = GetNeighborhood + (Not value(i, j + 1) And Neighbor.N_Bottom)
          
          'Diagonals
90        If diagonals Then
100           If chkLeft And chkTop Then GetNeighborhood = GetNeighborhood + (value(i - 1, j - 1) And Neighbor.N_TopLeft)
110           If chkLeft And chkBottom Then GetNeighborhood = GetNeighborhood + (value(i - 1, j + 1) And Neighbor.N_BottomLeft)
120           If chkRight And chkTop Then GetNeighborhood = GetNeighborhood + (value(i + 1, j - 1) And Neighbor.N_TopRight)
130           If chkRight And chkBottom Then GetNeighborhood = GetNeighborhood + (value(i + 1, j + 1) And Neighbor.N_BottomRight)
140       End If
              
End Function

Private Sub Class_Terminate()
          'Erase field
10        ReDim field(0)
          
End Sub
