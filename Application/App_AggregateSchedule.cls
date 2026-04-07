VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Builder     As App_EntityFilePathBuilder
    Repository  As Dom_IScheduleRepository
    Resolver    As Dom_ScheduleColumnResolver
    Factory     As Dom_ScheduleAggregateFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_EntityFilePathBuilder, ByVal Repository As Dom_IScheduleRepository)
    Set This.Builder = Builder
    Set This.Repository = Repository
    Set This.Resolver = New Dom_ScheduleColumnResolver
    Set This.Factory = New Dom_ScheduleAggregateFactory
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_ScheduleAggregate
    With New Dom_SchoolYear
        Dim SchoolYear As Dom_SchoolYear
        Set SchoolYear = .Create(Dom_Util_Date.GetSchoolYear(TargetDate))
    End With
    Dim FilePath As String
    FilePath = This.Builder.Builed(Schedule, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchedule", "File path is invalid."
    Dim RawCSVRows() As String
    RawCSVRows = This.Repository.Load(FilePath)
    If App_Util_Array.IsEmpty(RawCSVRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchedule", "RawCSVRows is empty."
    Dim ColumnGroup As VBA.Collection
    Set ColumnGroup = This.Resolver.Resolve(RawCSVRows)
    If ColumnGroup.Count = 0 Then Err.Raise AppErrNothingItem, "App_AggregateSchedule", "ColumnGroup item is nothing."
    Set Execute = This.Factory.Create(ColumnGroup)
End Function
