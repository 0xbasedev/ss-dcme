VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim c_pasting As Boolean
'Dim undoch As changes



Public Property Get pasting() As Boolean
10        pasting = c_pasting
End Property
Public Property Let pasting(newpasting As Boolean)
10        c_pasting = newpasting
End Property

Function getTile(X As Integer, Y As Integer) As Integer
10        getTile = SharedVar.clipdata(X, Y)
End Function

Friend Sub SetClipboardData(tiles() As Integer, ByRef bitfield As boolArray, Bounds As area)
10        On Error GoTo SetClipboardData_Error

20        SharedVar.clipdata = tiles
          
30        Call SharedVar.clipBitField.CopyFromBitfield(bitfield)
      '    Set SharedVar.clipBitField = bitfield
          
40        SharedVar.clipBoundaries = Bounds
50        SharedVar.clipHasData = True
60        Call frmGeneral.UpdateToolBarButtons


70        On Error GoTo 0
80        Exit Sub

SetClipboardData_Error:
90        HandleError Err, "ClipB.SetClipboardData"
End Sub

Function GetClipboardData() As Integer()
10        GetClipboardData = SharedVar.clipdata
End Function

Friend Function GetClipboardBoundaries() As area
10        GetClipboardBoundaries = SharedVar.clipBoundaries
End Function

'Function GetClipboardBitField() As boolArray
'    Set GetClipboardBitField = SharedVar.clipBitField
'End Function

Sub Paste(ByRef sel As selection)
10        Call sel.SetSelectionData(SharedVar.clipdata, SharedVar.clipBitField, SharedVar.clipBoundaries)
20        Call frmGeneral.SetCurrentTool(T_selection)
30        frmGeneral.UpdateToolBarButtons

End Sub

Private Sub Class_Initialize()
10        c_pasting = False
          
20        Set SharedVar.clipBitField = New boolArray
          
End Sub

Private Sub Class_Terminate()
20        Set SharedVar.clipBitField = Nothing
End Sub
