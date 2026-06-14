VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_UserUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    ScheduleQS      As App_IScheduleQueryService
    ClassHourQS     As App_IClassHourQueryService
    SubjectQS       As App_ISubjectQueryService
    PeriodQS        As App_IPeriodQueryService
    EnrollmentQS    As App_IEnrollmentQueryService
    MainStreamQS    As App_IMainStreamQueryService
    SpecialStreamQS As App_ISpecialStreamQueryService
    Base            As App_BaseUseCase
    PeriodPre       As Pre_DailyPeriodPresenter
    SchedulePre     As Pre_DailySchedulePresenter
End Type

Private This As Member

Friend Sub Inject(ByVal ScheduleQS As App_IScheduleQueryService, ByVal ClassHourQS As App_IClassHourQueryService, ByVal SubjectQS As App_ISubjectQueryService, ByVal PeriodQS As App_IPeriodQueryService, ByVal EnrollmentQS As App_IEnrollmentQueryService, ByVal MainStreamQS As App_IMainStreamQueryService, ByVal SpecialStreamQS As App_ISpecialStreamQueryService, ByVal Base As App_BaseUseCase, ByVal PeriodPre As Pre_DailyPeriodPresenter, ByVal SchedulePre As Pre_DailySchedulePresenter)
    Set This.ScheduleQS = ScheduleQS
    Set This.ClassHourQS = ClassHourQS
    Set This.SubjectQS = SubjectQS
    Set This.PeriodQS = PeriodQS
    Set This.EnrollmentQS = EnrollmentQS
    Set This.MainStreamQS = MainStreamQS
    Set This.SpecialStreamQS = SpecialStreamQS
    Set This.Base = Base
    Set This.PeriodPre = PeriodPre
    Set This.SchedulePre = SchedulePre
End Sub

Public Function CreateImportDailyPeriodUseCase() As App_ImportDailyPeriodUseCase
    Dim Result As App_ImportDailyPeriodUseCase
    Set Result = New App_ImportDailyPeriodUseCase
    Result.Inject This.ClassHourQS, This.MainStreamQS, This.Base, This.PeriodPre
    Set CreateImportDailyPeriodUseCase = Result
End Function

Public Function CreateImportDailyScheduleUseCase() As App_ImportDailyScheduleUseCase
    Dim Result As App_ImportDailyScheduleUseCase
    Set Result = New App_ImportDailyScheduleUseCase
    Result.Inject This.ScheduleQS, This.Base, This.SchedulePre
    Set CreateImportDailyScheduleUseCase = Result
End Function
