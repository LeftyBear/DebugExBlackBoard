VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_IScheduleQueryService
Private Type Member
    Persistence As Inf_SchedulePersistence
    DateIndex   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_SchedulePersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded(ByVal TargetDate As Date)
    If This.DateIndex Is Nothing Then
        Set This.DateIndex = New Scripting.Dictionary
    Else
        If This.DateIndex.Exists(CStr(TargetDate)) Then Exit Sub
    End If
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    Dim Rows As Inf_ScheduleRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    BuildDateIndex Rows
End Sub

Private Sub BuildDateIndex(ByVal Rows As Inf_ScheduleRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim TargetDate As Date
        TargetDate = Rows.Item(i).NormDate
        If 0 < TargetDate Then
            Dim DailyRows As Inf_ScheduleRows
            Set DailyRows = New Inf_ScheduleRows
            This.DateIndex.Add CStr(TargetDate), DailyRows
        End If
        DailyRows.Add Rows.Item(i)
    Next
End Sub

Private Sub RebuildDateIndex(ByVal Rows As Inf_ScheduleRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim TargetDate As Date
        TargetDate = Rows.Item(i).NormDate
        If 0 < TargetDate Then
            Dim DailyRows As Inf_ScheduleRows
            Set DailyRows = This.DateIndex.Item(CStr(TargetDate))
        End If
        Dim RowIndex As Long
        RowIndex = ((i - 1) Mod Rows.Count) + 1
        Set DailyRows.Item(RowIndex) = Rows.Item(i)
    Next
End Sub

Private Function App_IScheduleQueryService_GetByDate(ByVal TargetDate As Date) As App_ScheduleReadModels
    EnsureLoaded TargetDate
    Dim DailyRows As Inf_ScheduleRows
    Set DailyRows = This.DateIndex.Item(CStr(TargetDate))
    Dim EntityFactory As Inf_DailyScheduleFactory
    Set EntityFactory = New Inf_DailyScheduleFactory
    Dim DailyEntries As Dom_DailySchedules
    Set DailyEntries = EntityFactory.Create(DailyRows)
    Dim ModelFactory As Inf_DailyScheduleModelFactory
    Set ModelFactory = New Inf_DailyScheduleModelFactory
    Set App_IScheduleQueryService_GetByDate = ModelFactory.Create(DailyEntries)
End Function

Private Sub App_IScheduleQueryService_Rebuild(ByVal TargetDate As Date)
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    Dim Rows As Inf_ScheduleRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    RebuildDateIndex Rows
End Sub
