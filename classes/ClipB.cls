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
    pasting = c_pasting
End Property
Public Property Let pasting(newpasting As Boolean)
    c_pasting = newpasting
End Property

Function getTile(X As Integer, Y As Integer) As Integer
    getTile = SharedVar.clipdata(X, Y)
End Function

Friend Sub SetClipboardData(tiles() As Integer, ByRef bitfield As boolArray, Bounds As area)
    On Error GoTo SetClipboardData_Error

    SharedVar.clipdata = tiles
    
    Call SharedVar.clipBitField.CopyFromBitfield(bitfield)
'    Set SharedVar.clipBitField = bitfield
    
    SharedVar.clipBoundaries = Bounds
    SharedVar.clipHasData = True
    Call frmGeneral.UpdateToolBarButtons


    On Error GoTo 0
    Exit Sub

SetClipboardData_Error:
    HandleError Err, "ClipB.SetClipboardData"
End Sub

Function GetClipboardData() As Integer()
    GetClipboardData = SharedVar.clipdata
End Function

Friend Function GetClipboardBoundaries() As area
    GetClipboardBoundaries = SharedVar.clipBoundaries
End Function

'Function GetClipboardBitField() As boolArray
'    Set GetClipboardBitField = SharedVar.clipBitField
'End Function

Sub Paste(ByRef sel As selection)
    Call sel.SetSelectionData(SharedVar.clipdata, SharedVar.clipBitField, SharedVar.clipBoundaries)
    Call frmGeneral.SetCurrentTool(T_selection)
    frmGeneral.UpdateToolBarButtons

End Sub

Private Sub Class_Initialize()
    c_pasting = False
    
    Set SharedVar.clipBitField = New boolArray
    
End Sub

Private Sub Class_Terminate()
    Set SharedVar.clipBitField = Nothing
End Sub
