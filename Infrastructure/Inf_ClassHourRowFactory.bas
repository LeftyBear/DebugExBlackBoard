Attribute VB_Name = "Inf_ClassHourRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Column As Inf_ClassHourColumn, ByVal RawText As String) As Inf_ClassHourRow
    Dim Result As Inf_ClassHourRow
    Set Result = New Inf_ClassHourRow
    If Column.RawDate <> vbNullString Then
        Result.NormDate = Inf_DateUtility.NormalizeToDate(RawText)
    ElseIf Column.RawID <> vbNullString Then
        Dim Parts() As String
        Parts = VBA.Split(Column.RawID, DELIMITER)
        If UBound(Parts) = 2 Then
            Result.RawID = Column.RawID
            Dim Period As New Inf_PeriodMapper
            If 0 < VBA.InStr(1, Parts(1), Period.Key) Then
                Result.Mark = RawText
            Else
                Result.Value = Inf_NumericUtility.ExtractNumber(RawText)
            End If
        End If
    End If
    Set Create = Result
End Function
