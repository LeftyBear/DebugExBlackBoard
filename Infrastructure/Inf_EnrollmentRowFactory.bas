Attribute VB_Name = "Inf_EnrollmentRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Column As Inf_EnrollmentColumn, ByVal RawText As String) As Inf_EnrollmentRow
    Dim Result As Inf_EnrollmentRow
    Set Result = New Inf_EnrollmentRow
    If Column.RawDate <> vbNullString Then
        Result.NormDate = Inf_DateUtility.NormalizeToDate(RawText)
    ElseIf Column.RawID <> vbNullString Then
        Dim Parts() As String
        Parts = VBA.Split(Column.RawID, DELIMITER)
        If UBound(Parts) = 2 Then
            Result.RawID = Column.RawID
            Result.Value = Inf_NumericUtility.ExtractNumber(RawText)
        End If
    ElseIf Column.RawTransfer <> vbNullString Then
        Result.RawTransfer = RawText
    ElseIf Column.RawRemarks <> vbNullString Then
        Result.Remarks = RawText
    End If
    Set Create = Result
End Function
