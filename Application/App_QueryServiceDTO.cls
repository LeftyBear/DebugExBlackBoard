VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_QueryServiceDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule MoveFieldCloserToUsage, EncapsulatePublicField
'@Folder("Application.DTO")
Option Explicit

Public ScheduleQS       As App_IScheduleQueryService
Public SchoolEventQS    As App_ISchoolEventQueryService
Public ClassHourQS      As App_IClassHourQueryService
Public SubjectQS        As App_ISubjectQueryService
Public PeriodQS         As App_IPeriodQueryService
Public EnrollmentQS     As App_IEnrollmentQueryService
Public MainStreamQS     As App_IMainStreamQueryService
Public SpecialStreamQS  As App_ISpecialStreamQueryService
