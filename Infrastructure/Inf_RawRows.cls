VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_RawRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.DTOCollection")
Option Explicit
Private Type Member
    RawRows As VBA.Collection
    Rows    As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.RawRows = New VBA.Collection
    Set This.Rows = New VBA.Collection
End Sub

Public Sub AddRow(ByVal RawText As String)
    This.Rows.Add RawText
End Sub

Public Sub AddRows(ByVal Rows As VBA.Collection)
    This.RawRows.Add Rows
End Sub

Public Sub AddRowsInternal()
    This.RawRows.Add This.Rows
    Set This.Rows = New VBA.Collection
End Sub

Public Sub Remove(ByVal RowIndex As Long)
    If This.RawRows.Count = 0 Then Exit Sub
    If RowIndex < 1 Or This.RawRows.Count < RowIndex Then Exit Sub
    This.RawRows.Remove RowIndex
End Sub

Public Function RowsCount() As Long
    If This.RawRows.Count = 0 Then Exit Function
    RowsCount = This.RawRows.Count
End Function

Public Function ColumnsCount(ByVal RowIndex As Long) As Long
    If This.RawRows.Count = 0 Then Exit Function
    Dim Columns As VBA.Collection
    Set Columns = This.RawRows.Item(RowIndex)
    ColumnsCount = Columns.Count
End Function

Public Function GetHeaders() As VBA.Collection
    If This.RawRows.Count = 0 Then Exit Function
    Set GetHeaders = This.RawRows.Item(1)
End Function

Public Function GetHeader(ByVal ColumnIndex As Long) As String
    Dim Headers As VBA.Collection
    Set Headers = Me.GetHeaders
    If Headers Is Nothing Then Exit Function
    If ColumnIndex < 1 Or Headers.Count < ColumnIndex Then Exit Function
    GetHeader = Headers.Item(ColumnIndex)
End Function

Public Function GetRow(ByVal RowIndex As Long, ByVal ColumnIndex As Long) As String
    If This.RawRows.Count = 0 Then Exit Function
    If RowIndex < 1 Or This.RawRows.Count < RowIndex Then Exit Function
    Dim Columns As VBA.Collection
    Set Columns = This.RawRows.Item(RowIndex)
    If ColumnIndex < 1 Or Columns.Count < ColumnIndex Then Exit Function
    GetRow = Columns.Item(ColumnIndex)
End Function
