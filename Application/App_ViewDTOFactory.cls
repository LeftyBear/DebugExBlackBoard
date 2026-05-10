VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewDTOFactory"
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

Private Sub Class_Initialize()
    Set This.ClassHourFormatter = New App_ClassHourFormatter
    Set This.EnrollmentFormatter = New App_EnrollmentFormatter
    Set This.TimeTableFormatter = New App_TimeTableFormatter
End Sub
'
'Public Sub Initialize(ByVal ClassHourFormatter As App_ClassHourFormatter, ByVal EnrollmentFormatter As App_EnrollmentFormatter, ByVal TimeTableFormatter As App_TimeTableFormatter)
'    Set This.ClassHourFormatter = ClassHourFormatter
'    Set This.EnrollmentFormatter = EnrollmentFormatter
'    Set This.TimeTableFormatter = TimeTableFormatter
'End Sub

Public Function Create(ByVal Index As Long, ByVal ClassHour As VBA.Collection, ByVal Enrollment As VBA.Collection, ByVal Schedule As VBA.Collection, ByVal Structure As Dom_SchoolStructure) As App_ViewDTO
'    Dim BeginDate As Date
'    BeginDate = App_DateUtility.GetMonday(InitialDate.Value)
'    Dim BeginIndex As Long
'    BeginIndex = App_DateUtility.DiffApr1st(BeginDate)
'    Dim EndDate As Date
'    EndDate = App_DateUtility.GetSunday(InitialDate.Value)
'    Dim EndIndex As Long
'    EndIndex = App_DateUtility.DiffApr1st(EndDate)
    Dim ViewDTO As App_ViewDTO
    Set ViewDTO = New App_ViewDTO
    With ViewDTO
        Set .Structure = Structure
        Set .Schedule = Schedule.Item(Index)
        .ClassHourPlanTable = This.ClassHourFormatter.Format(Plan, ClassHour, Index, Index, Structure)
        .ClassHourExecutionTable = This.ClassHourFormatter.Format(Execution, ClassHour, Index, Index, Structure)
        .TimeTablePlanTable = This.TimeTableFormatter.Format(Plan, ClassHour.Item(Index), Structure)
        .TimeTableExecutionTable = This.TimeTableFormatter.Format(Execution, ClassHour.Item(Index), Structure)
        .EnrollmentTable = This.EnrollmentFormatter.Format(Enrollment.Item(Index), Structure)
    End With
    Set Create = ViewDTO
End Function
