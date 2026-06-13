Attribute VB_Name = "Inf_ClassHourRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawDate As String, ByVal Column As Inf_ClassHourColumn, ByVal RawText As String) As Inf_ClassHourRow
    Dim Result As Inf_ClassHourRow
    Set Result = New Inf_ClassHourRow
    Result.NormDate = Inf_DateUtility.NormalizeToDate(RawDate)
    If Column.Name <> vbNullString Then
        Dim Parts() As String
        Parts = VBA.Split(Column.Name, DELIMITER)
        If UBound(Parts) = 2 Then
            Result.RawID = Column.Name
            Result.RawValue = RawText
        End If
    End If
    Set Create = Result
End Function
