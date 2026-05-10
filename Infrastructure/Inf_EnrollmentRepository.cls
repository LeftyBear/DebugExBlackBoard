VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit
Implements Dom_IEnrollmentRepository
Private Type Member
    Persistence As Inf_EnrollmentPersistence
    AggList     As Scripting.Dictionary
    ChangedDate As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_EnrollmentPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded(ByVal SchoolYear As Long)
    If Not This.AggList Is Nothing Then Exit Sub
    Set This.ChangedDate = New VBA.Collection
    Dim Rows As Inf_EnrollmentRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    Dim Factory As Inf_YearlyEnrollmentFactory
    Set Factory = New Inf_YearlyEnrollmentFactory
    Dim Aggregates As Dom_YearlyEnrollment
    Set Aggregates = Factory.Create(Rows)
    Set This.AggList = BuiledDateIndex(Aggregates)
End Sub

Private Function BuiledDateIndex(ByVal Aggregates As Dom_YearlyEnrollment) As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    Dim Index As Long
    For Index = 1 To Aggregates.Count
        Dim Item As Dom_DailyEnrollment
        Set Item = Aggregates.Item(Index)
        Result.Add Item.DateID.ToKey, Item
    Next
    Set BuiledDateIndex = Result
End Function

Private Function Dom_IEnrollmentRepository_FindByDate(ByVal TargetDate As Date) As Dom_DailyEnrollment
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    EnsureLoaded SchoolYear
    If This.AggList.Exists(CStr(TargetDate)) Then
        Set Dom_IEnrollmentRepository_FindByDate = This.AggList.Item(CStr(TargetDate))
    End If
End Function

Private Sub Dom_IEnrollmentRepository_Replace(ByVal TargetDate As Date, ByVal NewAggregate As Dom_DailyEnrollment)
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    EnsureLoaded SchoolYear
    Set This.AggList.Item(CStr(TargetDate)) = NewAggregate
    This.ChangedDate.Add TargetDate
End Sub

Private Sub Dom_IEnrollmentRepository_Save(ByVal SchoolYear As Long)
    EnsureLoaded SchoolYear
    Dim Rows As Inf_EnrollmentRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    ReplaceByDaily Rows
    This.Persistence.WriteAll SchoolYear, Rows
End Sub

Private Sub ReplaceByDaily(ByRef Rows As Inf_EnrollmentRows)
    Dim i As Long
    For i = 1 To This.ChangedDate.Count
        Dim TargetDate As Date
        TargetDate = This.ChangedDate.Item(i)
        Dim Daily As Dom_DailyEnrollment
        Set Daily = This.AggList.Item(CStr(TargetDate))
        ReplaceRow Rows, Daily
    Next
End Sub

Private Sub ReplaceRow(ByRef Rows As Inf_EnrollmentRows, ByVal Daily As Dom_DailyEnrollment)
    Dim TargetDate As Date
    TargetDate = Daily.DateID.Value
    Dim DateCount As Long
    DateCount = Inf_DateUtility.DiffApr1st(TargetDate)
    Dim RowCount As Long
    RowCount = GetRowCount(Rows)
    Dim BeginIndex As Long
    BeginIndex = (DateCount - 1) * RowCount + 1
    Dim EndIndex As Long
    EndIndex = (DateCount - 1) * RowCount + RowCount
    Dim i As Long
    For i = BeginIndex To EndIndex
        Dim Row As Inf_EnrollmentRow
        Set Row = Rows.Item(i)
        Dim j As Long
        For j = 1 To Daily.Count
            Dim Entry As Dom_Enrollment
            Set Entry = Daily.Item(j)
            If Entry.ID.ToKey = Row.RawID Then
                Row.Value = Entry.Value
            ElseIf Entry.HasTransfer Then
                Row.Transfer = CStr(Entry.Transfer)
            ElseIf Entry.HasRemarks Then
                Row.Remarks = Entry.Remarks
            End If
        Next
    Next
End Sub

Private Function GetRowCount(ByVal Rows As Inf_EnrollmentRows) As Long
    Dim CurrentDate As Date
    CurrentDate = Rows.Item(1).NormDate
    Dim i As Long
    For i = 1 To Rows.Count
        If Rows.Item(i).NormDate <> CurrentDate Then Exit Function
        GetRowCount = GetRowCount + 1
    Next
End Function
