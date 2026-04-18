VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_ViewModelBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.ViewModel")
Option Explicit

Private Type Member
    EnrollmentFormatter As App_ViewEnrollmentFormatter
    ClassHourFormatter  As App_ViewClassHourFormatter
    TimeTableFormatter  As App_ViewTimeTableFormatter
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.EnrollmentFormatter = New App_ViewEnrollmentFormatter
    Set This.ClassHourFormatter = New App_ViewClassHourFormatter
    Set This.TimeTableFormatter = New App_ViewTimeTableFormatter
End Sub

Public Function Build(ByVal Schedule As Dom_ScheduleYearAggregate, ByVal Enrollment As Dom_EnrollmentYearAggregate, ByVal ClassHour As Dom_ClassHourYearAggregate, ByVal Structure As Dom_SchoolStructure, ByVal TargetDate As Dom_DateID) As Pre_ViewModel
    Dim ViewModel As Pre_ViewModel
    Set ViewModel = New Pre_ViewModel
    With ViewModel
        Set .Structure = Structure
        Set .Schedule = Schedule.GetByDate(TargetDate)
        .EnrollmentTable = This.EnrollmentFormatter.Format(Enrollment, TargetDate, Structure)
        .ClassHourPlanTable = This.ClassHourFormatter.Format(Plan, ClassHour, TargetDate, Structure)
        .ClassHourExecutionTable = This.ClassHourFormatter.Format(Execution, ClassHour, TargetDate, Structure)
        .TimeTablePlanTable = This.TimeTableFormatter.Format(Plan, ClassHour, TargetDate, Structure)
        .TimeTableExecutionTable = This.TimeTableFormatter.Format(Execution, ClassHour, TargetDate, Structure)
    End With
    Set Build = ViewModel
End Function
