Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit

Public Sub Run()
    Dim LimitValueUseCase As App_AggregateLimitValueUseCase
    Set LimitValueUseCase = App_UseCaseFactory.CreateAggregateLimitValueUseCase
    Dim LimitValue As Dom_EntityAggregate
    Set LimitValue = LimitValueUseCase.Execute
    Dim TotalizationUseCase As App_AggregateSubjectUseCase
    Set TotalizationUseCase = App_UseCaseFactory.CreateAggregateSubjectUseCase
    Dim Subject As Dom_EntityAggregate
    Set Subject = TotalizationUseCase.Execute
'    Dim SchoolStructure As Dom_SchoolStructure
'    Set SchoolStructure = New Dom_SchoolStructure
'    SchoolStructure.Initialize
    Dim EnrollmentUseCase As App_AggregateEnrollmentUseCase
    Set EnrollmentUseCase = App_UseCaseFactory.CreateAggregateEnrollmentUseCase
    Dim Enrollment As Dom_EntityAggregate
    Set Enrollment = EnrollmentUseCase.Execute(Date)
End Sub
