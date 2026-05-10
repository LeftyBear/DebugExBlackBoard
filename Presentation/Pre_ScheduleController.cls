VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_ScheduleController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Controller")
Option Explicit

Private Type Member
    Presenter As Pre_SchedulePresenter
End Type

Private This As Member

Public Sub Initialize(ByVal Presenter As Pre_SchedulePresenter)
    Set This.Presenter = Presenter
End Sub

Public Sub ViewSchedule(ByVal Index As Long, ByVal ClassHour As Dom_IClassHourRepository, ByVal Enrollment As Dom_IEnrollmentRepository, ByVal Schedule As Dom_IScheduleRepository, ByVal Structure As Dom_SchoolStructure)
'    Dim ClassHourAgg As VBA.Collection
'    Set ClassHourAgg = ClassHour.GetAll
'    Dim EnrollmentAgg As VBA.Collection
'    Set EnrollmentAgg = Enrollment.GetAll
'    Dim ScheduleAgg As VBA.Collection
'    Set ScheduleAgg = Schedule.GetAll
'    This.Presenter.Show Index, ClassHourAgg, EnrollmentAgg, ScheduleAgg, Structure
End Sub
