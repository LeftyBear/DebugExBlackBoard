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
    DateIndex   As Scripting.Dictionary
    DateCache   As VBA.Collection
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_EnrollmentPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded(ByVal SchoolYear As Long)
    If Not This.DateIndex Is Nothing Then Exit Sub
    Set This.DateIndex = New Scripting.Dictionary
    Set This.DateCache = New VBA.Collection
    Dim Rows As Inf_EnrollmentRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    Dim RowList As Scripting.Dictionary
    Set RowList = GroupByDate(Rows)
    BuildDateIndex RowList
End Sub

Private Sub BuildDateIndex(ByVal RowList As Scripting.Dictionary)
    Dim Factory As Inf_DailyEnrollmentFactory
    Set Factory = New Inf_DailyEnrollmentFactory
    Dim Key As Variant
    For Each Key In RowList.Keys
        Dim Daily As Dom_DailyEnrollments
        Set Daily = Factory.Create(RowList.Item(CStr(Key)))
        If Not Daily Is Nothing Then
            This.DateIndex.Add Key, Daily
        End If
    Next
End Sub

Private Sub ReplaceRow(ByVal Rows As Inf_EnrollmentRows)
    Dim RowCount As Long
    RowCount = GetDailyRowCount(Rows)
    Dim i As Long
    For i = 1 To This.DateCache.Count
        Dim TargetDate As Date
        TargetDate = This.DateCache.Item(i)
        Dim DateIndex As Long
        DateIndex = Inf_DateUtility.DiffApr1st(TargetDate)
        Dim BeginIndex As Long
        BeginIndex = (DateIndex - 1) * RowCount + 1
        Dim EndIndex As Long
        EndIndex = BeginIndex + RowCount - 1
        Dim Daily As Dom_DailyEnrollments
        Set Daily = This.DateIndex.Item(CStr(TargetDate))
        ReplaceRowOfDaily BeginIndex, EndIndex, Rows, Daily
    Next
End Sub

Private Sub ReplaceRowOfDaily(ByVal BeginIndex As Long, ByVal EndIndex As Long, ByVal Rows As Inf_EnrollmentRows, ByVal Daily As Dom_DailyEnrollments)
    Dim i As Long
    For i = BeginIndex To EndIndex
        Dim Row As Inf_EnrollmentRow
        Set Row = Rows.Item(i)
        ReplaceRowOfDailyEnrollment Row, Daily
    Next
End Sub

Private Sub ReplaceRowOfDailyEnrollment(ByVal Row As Inf_EnrollmentRow, ByVal Daily As Dom_DailyEnrollments)
    Dim Mapper As Inf_EnrollmentIDMapper
    Set Mapper = New Inf_EnrollmentIDMapper
    Dim i As Long
    For i = 1 To Daily.Count
        Dim Entry As Dom_Enrollment
        Set Entry = Daily.Item(i)
        If Row.RawID = Mapper.ToRaw(Entry.ID) Then
            Row.Value = Entry.Value.Value
        ElseIf Entry.HasTransfer Then
            Row.RawTransfer = CStr(Entry.Transfer.IsTrue)
        ElseIf Entry.HasRemarks Then
            Row.Remarks = Entry.Remarks.Value
        End If
    Next
End Sub

Private Function GroupByDate(ByVal Rows As Inf_ClassHourRows) As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Key As String
        Key = CStr(Rows.Item(i).NormDate)
        If Not This.DateIndex.Exists(Key) Then
            Result.Add Key, New VBA.Collection
        End If
        Dim Items As VBA.Collection
        Set Items = This.DateIndex.Item(Key)
        Items.Add Rows.Item(i)
    Next
    Set GroupByDate = Result
End Function

Private Function GetDailyRowCount(ByVal Rows As Inf_EnrollmentRows) As Long
    Dim CurrentDate As Date
    CurrentDate = Rows.Item(1).NormDate
    Dim i As Long
    For i = 1 To Rows.Count
        If Rows.Item(i).NormDate <> CurrentDate Then Exit For
    Next
    GetDailyRowCount = i - 1
End Function

Private Function Dom_IEnrollmentRepository_FindByDate(ByVal TargetDate As Date) As Dom_DailyEnrollments
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    EnsureLoaded SchoolYear
    If This.DateIndex.Exists(CStr(TargetDate)) Then
        Set Dom_IEnrollmentRepository_FindByDate = This.DateIndex.Item(CStr(TargetDate))
    End If
End Function

Private Sub Dom_IEnrollmentRepository_Replace(ByVal TargetDate As Date, ByVal NewAggregate As Dom_DailyEnrollments)
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    EnsureLoaded SchoolYear
    Set This.DateIndex.Item(CStr(TargetDate)) = NewAggregate
    This.DateCache.Add TargetDate
End Sub

Private Sub Dom_IEnrollmentRepository_Save(ByVal SchoolYear As Long)
    EnsureLoaded SchoolYear
    Dim Rows As Inf_EnrollmentRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    ReplaceRow Rows
    This.Persistence.WriteAll SchoolYear, Rows
End Sub
