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
    QueryService As Inf_ScheduleQueryService
End Type

Private This As Member

Public Sub Inject(ByVal QueryService As Inf_ScheduleQueryService)
    Set This.QueryService = QueryService
End Sub
