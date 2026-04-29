VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewDTOBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    ClassHourFormatter  As App_ClassHourFormatter
    EnrollmentFormatter As App_EnrollmentFormatter
    TimeTableFormatter  As App_TimeTableFormatter
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourFormatter As App_ClassHourFormatter, ByVal EnrollmentFormatter As App_EnrollmentFormatter, ByVal TimeTableFormatter As App_TimeTableFormatter)
    Set This.ClassHourFormatter = ClassHourFormatter
    Set This.EnrollmentFormatter = EnrollmentFormatter
    Set This.TimeTableFormatter = TimeTableFormatter
End Sub

Public Function Build(ByVal Index As Long, ByVal ClassHour As Dom_ClassHourYearAggregate, ByVal Enrollment As Dom_EnrollmentYearAggregate, ByVal Schedule As Dom_ScheduleYearAggregate, ByVal Structure As Dom_SchoolStructure) As App_ViewDTO
    Dim ViewDTO As App_ViewDTO
    Set ViewDTO = New App_ViewDTO
    With ViewDTO
        Set .Structure = Structure
        Set .Schedule = Schedule.Item(Index)
        .ClassHourPlanTable = This.ClassHourFormatter.Format(Plan, ClassHour.Item(Index), Index, Index, Structure)
        .ClassHourExecutionTable = This.ClassHourFormatter.Format(Execution, ClassHour.Item(Index), Index, Index, Structure)
        .TimeTablePlanTable = This.TimeTableFormatter.Format(Plan, ClassHour.Item(Index), Structure)
        .TimeTableExecutionTable = This.TimeTableFormatter.Format(Execution, ClassHour.Item(Index), Structure)
        .EnrollmentTable = This.EnrollmentFormatter.Format(Enrollment.Item(Index), Structure)
    End With
    Set Build = ViewDTO
End Function
