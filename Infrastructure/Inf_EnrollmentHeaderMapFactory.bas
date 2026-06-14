Attribute VB_Name = "Inf_EnrollmentHeaderMapFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module
Private Const COLUMN_HIZUKE     As String = "ì˙ït"
Private Const COLUMN_TRANSFER   As String = "ç›ê–àŸìÆ"
Private Const COLUMN_REMARKS    As String = "ç›ê–îıçl"

Public Function Create(ByRef Header() As Variant) As Inf_EnrollmentHeaderMap
    Dim Result As Inf_EnrollmentHeaderMap
    Set Result = New Inf_EnrollmentHeaderMap
    Dim C As Long
    For C = LBound(Header) To UBound(Header)
        Dim Column As Inf_EnrollmentColumn
        Set Column = CreateColumn(Header(C))
        Result.Add CStr(C), Column
    Next
    Set Create = Result
End Function

Private Function CreateColumn(ByVal ColumnName As String) As Inf_EnrollmentColumn
    Dim Result As Inf_EnrollmentColumn
    Set Result = New Inf_EnrollmentColumn
    If 0 < VBA.InStr(1, ColumnName, COLUMN_HIZUKE) Then
        Result.RawDate = ColumnName
    ElseIf 0 < VBA.InStr(1, ColumnName, COLUMN_TRANSFER) Then
        Result.RawTransfer = ColumnName
    ElseIf 0 < VBA.InStr(1, ColumnName, COLUMN_REMARKS) Then
        Result.RawRemarks = ColumnName
    Else
        Result.RawID = ColumnName
    End If
    Set CreateColumn = Result
End Function
