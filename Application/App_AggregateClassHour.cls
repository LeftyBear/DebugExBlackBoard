VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Builder     As App_EntityFilePathBuilder
    Repository  As Dom_IClassHourRepository
    Resolver    As Dom_ClassHourColumnResolver
    Factory     As Dom_ClassHourAggregateFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_EntityFilePathBuilder, ByVal Repository As Dom_IClassHourRepository, ByVal Resolver As Dom_ClassHourColumnResolver)
    Set This.Builder = Builder
    Set This.Repository = Repository
    Set This.Resolver = Resolver
    Set This.Factory = New Dom_ClassHourAggregateFactory
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_ClassHourYearAggregate
    With New Dom_SchoolYear
        Dim SchoolYear As Dom_SchoolYear
        Set SchoolYear = .Create(Dom_Util_Date.GetSchoolYear(TargetDate))
    End With
    Dim FilePath As String
    FilePath = This.Builder.Builed(ClassHour, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateClassHour", "File path is invalid."
    Dim RawCSVRows() As String
    RawCSVRows = This.Repository.Load(SchoolYear)
    If App_Util_Array.IsEmpty(RawCSVRows) Then Err.Raise AppErrEmptyArray, "App_AggregateClassHour", "RawCSVRows is empty."
    Dim EntityGroups As VBA.Collection
    Set EntityGroups = This.Resolver.Resolve(RawCSVRows)
    If EntityGroups.Count = 0 Then Err.Raise AppErrNothingItem, "App_AggregateClassHour", "EntityGroups item is nothing."
    Set Execute = This.Factory.Create(EntityGroups)
End Function
