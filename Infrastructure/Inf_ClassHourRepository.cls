VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit
Implements Dom_IClassHourRepository
Private Type Member
    Persistence     As Inf_ClassHourPersistence
    ClassHourList   As Scripting.Dictionary
    DateCache       As VBA.Collection
End Type

Private This As Member

Friend Sub Inject(ByVal Persistence As Inf_ClassHourPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded(ByVal SchoolYear As Long)
    If Not This.ClassHourList Is Nothing Then Exit Sub
    Set This.ClassHourList = New Scripting.Dictionary
    Set This.DateCache = New VBA.Collection
    Dim Rows As Inf_ClassHourRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    Dim RowList As Scripting.Dictionary
    Set RowList = GroupByDate(Rows)
    BuildClassHourList RowList
End Sub

Private Sub BuildClassHourList(ByVal RowList As Scripting.Dictionary)
    Dim Factory As Inf_DailyClassHourFactory
    Set Factory = New Inf_DailyClassHourFactory
    Dim Key As Variant
    For Each Key In RowList.Keys
        Dim Daily As Dom_DailyClassHours
        Set Daily = Factory.Create(RowList.Item(CStr(Key)))
        If Not Daily Is Nothing Then
            This.ClassHourList.Add Key, Daily
        End If
    Next
End Sub

Private Sub ReplaceRow(ByVal Rows As Inf_ClassHourRows)
    Dim RowCount As Long
    RowCount = GetDailyRowCount(Rows)
    Dim i As Long
    For i = 1 To This.DateCache.Count
        Dim TargetDate As Date
        TargetDate = This.DateCache.Item(i)
        Dim ClassHourList As Long
        ClassHourList = Inf_DateUtility.DiffApr1st(TargetDate)
        Dim BeginIndex As Long
        BeginIndex = (ClassHourList - 1) * RowCount + 1
        Dim EndIndex As Long
        EndIndex = BeginIndex + RowCount - 1
        Dim Daily As Dom_DailyClassHours
        Set Daily = This.ClassHourList.Item(CStr(TargetDate))
        ReplaceRowOfDaily BeginIndex, EndIndex, Rows, Daily
    Next
End Sub

Private Sub ReplaceRowOfDaily(ByVal BeginIndex As Long, ByVal EndIndex As Long, ByVal Rows As Inf_ClassHourRows, ByVal Daily As Dom_DailyClassHours)
    Dim Mapper As Inf_PeriodMapper
    Set Mapper = New Inf_PeriodMapper
    Dim i As Long
    For i = BeginIndex To EndIndex
        Dim Row As Inf_ClassHourRow
        Set Row = Rows.Item(i)
        If Mapper.ToDomain(Row.RawID) Is Nothing Then
            ReplaceRowOfDailyClassHour Row, Daily
        Else
            ReplaceRowOfDailyTimeTable Row, Daily
        End If
    Next
End Sub

Private Sub ReplaceRowOfDailyTimeTable(ByVal Row As Inf_ClassHourRow, ByVal Daily As Dom_DailyClassHours)
    Dim Mapper As Inf_TimeTableIDMapper
    Set Mapper = New Inf_TimeTableIDMapper
    Dim i As Long
    For i = 1 To Daily.TimeTablesCount
        Dim Entry As Dom_TimeTable
        Set Entry = Daily.TimeTablesItem(i)
        If Row.RawID = Mapper.ToRaw(Entry.ID) Then
            Row.RawValue = Entry.Value.Value
        End If
    Next
End Sub

Private Sub ReplaceRowOfDailyClassHour(ByVal Row As Inf_ClassHourRow, ByVal Daily As Dom_DailyClassHours)
    Dim Mapper As Inf_ClassHourIDMapper
    Set Mapper = New Inf_ClassHourIDMapper
    Dim i As Long
    For i = 1 To Daily.ClassHoursCount
        Dim Entry As Dom_ClassHour
        Set Entry = Daily.ClassHoursItem(i)
        If Row.RawID = Mapper.ToRaw(Entry.ID) Then
            Row.RawValue = Entry.Value.Value
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
        If Not This.ClassHourList.Exists(Key) Then
            Result.Add Key, New VBA.Collection
        End If
        Dim Items As VBA.Collection
        Set Items = This.ClassHourList.Item(Key)
        Items.Add Rows.Item(i)
    Next
    Set GroupByDate = Result
End Function

Private Function GetDailyRowCount(ByVal Rows As Inf_ClassHourRows) As Long
    Dim CurrentDate As Date
    CurrentDate = Rows.Item(1).NormDate
    Dim i As Long
    For i = 1 To Rows.Count
        If Rows.Item(i).NormDate <> CurrentDate Then Exit For
    Next
    GetDailyRowCount = i - 1
End Function

Private Function Dom_IClassHourRepository_FindByDate(ByVal TargetDate As Date) As Dom_DailyClassHours
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    EnsureLoaded SchoolYear
    If This.ClassHourList.Exists(CStr(TargetDate)) Then
        Set Dom_IClassHourRepository_FindByDate = This.ClassHourList.Item(CStr(TargetDate))
    End If
End Function

Private Sub Dom_IClassHourRepository_Replace(ByVal TargetDate As Date, ByVal NewAggregate As Dom_DailyClassHours)
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    EnsureLoaded SchoolYear
    Set This.ClassHourList.Item(CStr(TargetDate)) = NewAggregate
    This.DateCache.Add TargetDate
End Sub

Private Sub Dom_IClassHourRepository_Save(ByVal SchoolYear As Long)
    EnsureLoaded SchoolYear
    Dim Rows As Inf_ClassHourRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    ReplaceRow Rows
    This.Persistence.WriteAll SchoolYear, Rows
End Sub
