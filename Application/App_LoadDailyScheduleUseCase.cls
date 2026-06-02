VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadDailyScheduleUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    QueryService As App_IScheduleQueryService
End Type

Private This As Member

Friend Sub Inject(ByVal QueryService As App_IScheduleQueryService)
    Set This.QueryService = QueryService
End Sub

Public Function Execute(ByVal TargetDate As Date) As App_ScheduleReadModels
    Set Execute = This.QueryService.GetByDate(TargetDate)
End Function
