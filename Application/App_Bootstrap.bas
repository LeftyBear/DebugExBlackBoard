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
    'UseCase-------------------------------------------------------------------------------------
    Dim EnrollmentUseCase As App_AggregateEnrollmentUseCase
    Set EnrollmentUseCase = App_UseCaseFactory.CreateAggregateEnrollmentUseCase(SchoolConfig)
    Dim ClassHourUseCase As App_AggregateClassHourUseCase
    Set ClassHourUseCase = App_UseCaseFactory.CreateAggregateClassHourUseCase(SchoolConfig)
    'Aggregate-----------------------------------------------------------------------------------
    Dim Enrollment As Dom_EnrollmentYearAggregate
    Set Enrollment = EnrollmentUseCase.Execute(Date)
    Dim ClassHour As Dom_ClassHourYearAggregate
    Set ClassHour = ClassHourUseCase.Execute(Date)
    'View----------------------------------------------------------------------------------------
    Dim EnrollmentFormatter As App_ViewEnrollmentFormatter
    Set EnrollmentFormatter = New App_ViewEnrollmentFormatter
    Dim EnrollmentTable() As Variant
    EnrollmentTable = EnrollmentFormatter.Format(Enrollment.GetAggregate(Date), SchoolConfig)
    Dim ClassHourFormatter As App_ViewClassHourFormatter
    Set ClassHourFormatter = New App_ViewClassHourFormatter
    Dim ClassHourPlanTable() As Variant
    ClassHourPlanTable = ClassHourFormatter.Format(Plan, ClassHour.GetAggregate(Date), SchoolConfig)
    Dim ClassHourExecutionTable() As Variant
    ClassHourExecutionTable = ClassHourFormatter.Format(Execution, ClassHour.GetAggregate(Date), SchoolConfig)
    Dim TimeTableFormatter As App_ViewTimeTableFormatter
    Set TimeTableFormatter = New App_ViewTimeTableFormatter
    Dim TimeTablePlanTable() As Variant
    TimeTablePlanTable = TimeTableFormatter.Format(Plan, ClassHour.GetAggregate(Date), SchoolConfig)
    Dim TimeTableExecutionTable() As Variant
    TimeTableExecutionTable = TimeTableFormatter.Format(Execution, ClassHour.GetAggregate(Date), SchoolConfig)
End Sub
