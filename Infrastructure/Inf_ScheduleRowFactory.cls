VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleRowFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Column As Inf_ScheduleColumn, ByVal RawText As String) As Inf_ScheduleRow
    Dim Result As Inf_ScheduleRow
    Set Result = New Inf_ScheduleRow
    If Column.RawDate <> vbNullString Then
        Result.NormDate = Inf_DateUtility.NormalizeToDate(RawText)
    ElseIf Column.RawID <> vbNullString Then
        Dim Parts() As String
        Parts = VBA.Split(Column.RawID, charUnderScore)
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
