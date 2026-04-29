VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ScheduleAggregator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Reader      As App_IScheduleReader
    Builder     As App_EntityFilePathBuilder
    Resolver    As App_ScheduleRowResolver
    Factory     As App_ScheduleYearAggFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As App_IScheduleReader, ByVal Builder As App_EntityFilePathBuilder, ByVal Resolver As App_ScheduleRowResolver, ByVal Factory As App_ScheduleYearAggFactory)
    Set This.Reader = Reader
    Set This.Builder = Builder
    Set This.Resolver = Resolver
    Set This.Factory = Factory
End Sub

Public Function Execute(ByVal SchoolYear As Dom_SchoolYear, ByVal Structure As Dom_SchoolStructure) As Dom_ScheduleYearAggregate
    Dim FilePath As String
    FilePath = This.Builder.Builed(Schedule, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim RawRows() As String
    RawRows = This.Reader.Load(FilePath)
    If App_ArrayUtility.IsEmpty(RawRows) Then Err.Raise AppErrEmptyArray, "App_ScheduleAggregator", "RawRows is empty."
    Dim Rows As VBA.Collection
    Set Rows = This.Resolver.Resolve(RawRows)
    If Rows.Count = 0 Then Err.Raise AppErrNothingItem, "App_ScheduleAggregator", "Rows item is nothing."
    Set Execute = This.Factory.Create(Rows)
End Function
