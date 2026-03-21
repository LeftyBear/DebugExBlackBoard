Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Sub Run()
    'SchoolStructure--------------------------------------------------------------------------------
    Dim SchoolConfigGenerater As App_GenerateSchoolStructure
    Set SchoolConfigGenerater = App_UseCaseFactory.CreateSchoolConfigGenerater
    Dim SchoolStructure As Dom_SchoolStructure
    Set SchoolStructure = SchoolConfigGenerater.Execute
    'UseCase-------------------------------------------------------------------------------------
    Dim EnrollmentUseCase As App_AggregateEnrollment
    Set EnrollmentUseCase = App_UseCaseFactory.CreateAggregateEnrollmentUseCase(SchoolStructure)
    Dim ClassHourUseCase As App_AggregateClassHour
    Set ClassHourUseCase = App_UseCaseFactory.CreateAggregateClassHourUseCase(SchoolStructure)
    'Aggregate-----------------------------------------------------------------------------------
    Dim Enrollment As Dom_EnrollmentYearAggregate
    Set Enrollment = EnrollmentUseCase.Execute(Date)
    Dim ClassHour As Dom_ClassHourYearAggregate
    Set ClassHour = ClassHourUseCase.Execute(Date)
    'View----------------------------------------------------------------------------------------
'    Dim EnrollmentFormatter As App_ViewEnrollmentFormatter
'    Set EnrollmentFormatter = New App_ViewEnrollmentFormatter
'    Dim EnrollmentTable() As Variant
'    EnrollmentTable = EnrollmentFormatter.Format(Enrollment.GetAggregate(Date), SchoolStructure)
'    Dim ClassHourFormatter As App_ViewClassHourFormatter
'    Set ClassHourFormatter = New App_ViewClassHourFormatter
'    Dim ClassHourPlanTable() As Variant
'    ClassHourPlanTable = ClassHourFormatter.Format(Plan, ClassHour.GetAggregate(Date), SchoolStructure)
'    Dim ClassHourExecutionTable() As Variant
'    ClassHourExecutionTable = ClassHourFormatter.Format(Execution, ClassHour.GetAggregate(Date), SchoolStructure)
'    Dim TimeTableFormatter As App_ViewTimeTableFormatter
'    Set TimeTableFormatter = New App_ViewTimeTableFormatter
'    Dim TimeTablePlanTable() As Variant
'    TimeTablePlanTable = TimeTableFormatter.Format(Plan, ClassHour.GetAggregate(Date), SchoolStructure)
'    Dim TimeTableExecutionTable() As Variant
'    TimeTableExecutionTable = TimeTableFormatter.Format(Execution, ClassHour.GetAggregate(Date), SchoolStructure)
End Sub
