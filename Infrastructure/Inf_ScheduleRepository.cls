VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_IScheduleRepository
Private Type Member
    Persistence As Inf_SchedulePersistence
    AggList     As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_SchedulePersistence)
    Set This.Persistence = Persistence
    Set This.AggList = New Scripting.Dictionary
End Sub

Private Sub Dom_IScheduleRepository_Add(ByVal Aggregate As Dom_DailySchedule)
    This.AggList.Add Aggregate.DateID.ToKey, Aggregate
End Sub

Private Function Dom_IScheduleRepository_FindByDate(ByVal TargetDate As Date) As Dom_DailySchedule
    If This.AggList.Exists(CStr(TargetDate)) Then
        Set Dom_IScheduleRepository_FindByDate = This.AggList.Item(CStr(TargetDate))
    End If
End Function

Private Function Dom_IScheduleRepository_GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Aggregate As Dom_DailySchedule
    For Each Aggregate In This.AggList.Items
        Result.Add Aggregate
    Next
    Set Dom_IScheduleRepository_GetAll = Result
End Function

Private Function Dom_IScheduleRepository_ReadAll(ByVal SchoolYear As Long) As VBA.Collection
    Set Dom_IScheduleRepository_ReadAll = This.Persistence.ReadAll(SchoolYear)
End Function

Private Function Dom_IScheduleRepository_Replace(ByVal NewAggregate As Dom_DailySchedule) As VBA.Collection
    Dim Rows As VBA.Collection
    Set Rows = Dom_IScheduleRepository_GetAll
    Dim i As Long
    For i = 1 To Rows.Count
        Dim OldAggregate As Dom_DailySchedule
        Set OldAggregate = Rows.Item(i)
        If OldAggregate.DateID.Equals(NewAggregate.DateID) Then
            Set Rows.Item(i) = NewAggregate
        End If
    Next
    Set Dom_IScheduleRepository_Replace = Rows
End Function
