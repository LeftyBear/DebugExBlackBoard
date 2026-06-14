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
    Base    As App_BaseUseCase
    QSDTO   As App_QueryServiceDTO
    PreDTO  As App_PresenterDTO
End Type

Private This As Member

Friend Sub Inject(ByVal Base As App_BaseUseCase, ByVal QSDTO As App_QueryServiceDTO, ByVal PreDTO As App_PresenterDTO)
    Set This.Base = Base
    Set This.QSDTO = QSDTO
    Set This.PreDTO = PreDTO
End Sub

Public Function CreateImportDailyPeriodUseCase() As App_ImportDailyPeriodUseCase
    Dim Result As App_ImportDailyPeriodUseCase
    Set Result = New App_ImportDailyPeriodUseCase
    Result.Inject This.QSDTO.ClassHourQS, This.QSDTO.MainStreamQS, This.Base, This.PreDTO.PeriodPre
    Set CreateImportDailyPeriodUseCase = Result
End Function

Public Function CreateImportDailyScheduleUseCase() As App_ImportDailyScheduleUseCase
    Dim Result As App_ImportDailyScheduleUseCase
    Set Result = New App_ImportDailyScheduleUseCase
    Result.Inject This.QSDTO.ScheduleQS, This.Base, This.PreDTO.SchedulePre
    Set CreateImportDailyScheduleUseCase = Result
End Function
