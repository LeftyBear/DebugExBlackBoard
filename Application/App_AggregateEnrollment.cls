VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Builder     As App_EntityFilePathBuilder
    Repository  As Dom_IEnrollmentRepository
    Resolver    As Dom_EnrollmentColumnResolver
    Factory     As Dom_EnrollmentAggregateFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_EntityFilePathBuilder, ByVal Repository As Dom_IEnrollmentRepository, ByVal Resolver As Dom_EnrollmentColumnResolver)
    Set This.Builder = Builder
    Set This.Repository = Repository
    Set This.Resolver = Resolver
    Set This.Factory = New Dom_EnrollmentAggregateFactory
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_EnrollmentYearAggregate
    With New Dom_SchoolYear
        Dim SchoolYear As Dom_SchoolYear
        Set SchoolYear = .Create(Dom_Util_Date.GetSchoolYear(TargetDate))
    End With
    Dim FilePath As String
    FilePath = This.Builder.Builed(Enrollment, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateEnrollment", "File path is invalid."
    Dim RawCSVRows() As String
    RawCSVRows = This.Repository.Load(FilePath)
    If App_Util_Array.IsEmpty(RawCSVRows) Then Err.Raise AppErrEmptyArray, "App_AggregateEnrollment", "RawCSVRows is empty."
    Dim EntityGroups As VBA.Collection
    Set EntityGroups = This.Resolver.Resolve(RawCSVRows)
    If EntityGroups.Count = 0 Then Err.Raise AppErrNothingItem, "App_AggregateEnrollment", "EntityGroups item is nothing."
    Set Execute = This.Factory.Create(EntityGroups)
End Function
