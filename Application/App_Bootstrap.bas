Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Sub Run()
    'SchoolConfig--------------------------------------------------------------------------------
    Dim SchoolConfigGenerater As App_GenerateSchoolConfigUseCase
    Set SchoolConfigGenerater = App_UseCaseFactory.CreateSchoolConfigGenerater
    Dim SchoolConfig As Dom_SchoolConfig
    Set SchoolConfig = SchoolConfigGenerater.Execute
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
