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
    Logger          As App_ILogPersistence
    ClassHourQS     As App_IClassHourQueryService
    ScheduleQS      As App_IScheduleQueryService
    MainStreamQS    As App_IMainStreamQueryService
    Base            As App_BaseUseCase
    PeriodPre       As Pre_DailyPeriodPresenter
    SchedulePre     As Pre_DailySchedulePresenter
End Type

Private This As Member

Friend Sub Inject(ByVal Logger As App_ILogPersistence, ByVal ClassHourQS As App_IClassHourQueryService, ByVal ScheduleQS As App_IScheduleQueryService, ByVal MainStreamQS As App_IMainStreamQueryService, ByVal Base As App_BaseUseCase, ByVal PeriodPre As Pre_DailyPeriodPresenter, ByVal SchedulePre As Pre_DailySchedulePresenter)
    Set This.Logger = Logger
    Set This.ClassHourQS = ClassHourQS
    Set This.ScheduleQS = ScheduleQS
    Set This.MainStreamQS = MainStreamQS
    Set This.Base = Base
    Set This.PeriodPre = PeriodPre
    Set This.SchedulePre = SchedulePre
End Sub

Public Function CreateImportDailyPeriodUseCase() As App_ImportDailyPeriodUseCase
    Dim Result As App_ImportDailyPeriodUseCase
    Set Result = New App_ImportDailyPeriodUseCase
    Result.Inject This.Logger, This.ClassHourQS, This.MainStreamQS, This.Base, This.PeriodPre
    Set CreateImportDailyPeriodUseCase = Result
End Function

Public Function CreateImportDailyScheduleUseCase() As App_ImportDailyScheduleUseCase
    Dim Result As App_ImportDailyScheduleUseCase
    Set Result = New App_ImportDailyScheduleUseCase
    Result.Inject This.Logger, This.ScheduleQS, This.Base, This.SchedulePre
    Set CreateImportDailyScheduleUseCase = Result
End Function
