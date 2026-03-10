Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Sub Run()
    'SchoolConfig--------------------------------------------------------------------------------
    Dim LimitValueUseCase As App_AggregateLimitValueUseCase
    Set LimitValueUseCase = App_UseCaseFactory.CreateAggregateLimitValueUseCase
    Dim LimitValue As Dom_EntityAggregate
    Set LimitValue = LimitValueUseCase.Execute
    Dim SubjectUseCase As App_AggregateSubjectUseCase
    Set SubjectUseCase = App_UseCaseFactory.CreateAggregateSubjectUseCase
    Dim Subject As Dom_EntityAggregate
    Set Subject = SubjectUseCase.Execute
    Dim ConfigUseCase As App_GenerateConfigUseCase
    Set ConfigUseCase = New App_GenerateConfigUseCase
    ConfigUseCase.Initialize LimitValue, Subject
    Dim SchoolConfig As Dom_SchoolConfig
    Set SchoolConfig = ConfigUseCase.Execute
    'Enrollment----------------------------------------------------------------------------------
    Dim EnrollmentUseCase As App_AggregateEnrollmentUseCase
    Set EnrollmentUseCase = App_UseCaseFactory.CreateAggregateEnrollmentUseCase(SchoolConfig)
    Dim Enrollment As Dom_EnrollmentAggregate
    Set Enrollment = EnrollmentUseCase.Execute(Date)
    Dim EnrollmentTable() As Variant
    Dim EnrollmentFormatter As App_ViewEnrollmentFormatter
    Set EnrollmentFormatter = New App_ViewEnrollmentFormatter
    Dim DateIndexCalculator As Dom_DateIndexCalculator
    Set DateIndexCalculator = New Dom_DateIndexCalculator
    Dim DateIndex As Long
    DateIndex = DateIndexCalculator.Calculate(Date)
    EnrollmentTable = EnrollmentFormatter.Format(Enrollment.GetRecord(DateIndex), SchoolConfig)
End Sub
