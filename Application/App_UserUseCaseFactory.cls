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

Public Function CreateLoadPeriodUseCase() As App_LoadPeriodUseCase
    Dim Result As App_LoadPeriodUseCase
    Set Result = New App_LoadPeriodUseCase
    Result.Inject This.QSDTO.ClassHourQS, This.QSDTO.MainStreamQS, This.Base, This.PreDTO.PeriodPre
    Set CreateLoadPeriodUseCase = Result
End Function

Public Function CreateLoadScheduleUseCase() As App_LoadScheduleUseCase
    Dim Result As App_LoadScheduleUseCase
    Set Result = New App_LoadScheduleUseCase
    Result.Inject This.QSDTO.ScheduleQS, This.Base, This.PreDTO.SchedulePre
    Set CreateLoadScheduleUseCase = Result
End Function
