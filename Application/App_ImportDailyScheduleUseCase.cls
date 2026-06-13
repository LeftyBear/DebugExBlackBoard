VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ImportDailyScheduleUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit
Implements App_IUseCaseCallback
Private Type Member
    Logger          As App_ILogPersistence
    QueryService    As App_IScheduleQueryService
    Base            As App_BaseUseCase
    Presenter       As Pre_DailySchedulePresenter
    TargetDate      As Date
End Type

Private This As Member

Friend Sub Inject(ByVal Logger As App_ILogPersistence, ByVal QueryService As Inf_ScheduleQueryService, ByVal Base As App_BaseUseCase, ByVal Presenter As Pre_DailySchedulePresenter)
    Set This.Logger = Logger
    Set This.QueryService = QueryService
    Set This.Base = Base
    Set This.Presenter = Presenter
End Sub

Public Sub Execute(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
    This.Base.Execute Me
End Sub

Private Sub App_IUseCaseCallback_ExecuteCore()
    Dim ViewModels As App_ScheduleReadModels
    Set ViewModels = This.QueryService.GetByDate(This.TargetDate)
    This.Presenter.Present ViewModels
End Sub

Private Sub App_IUseCaseCallback_LogSystemError(ByVal Error As VBA.ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  Source: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.Log Message
End Sub

Private Sub App_IUseCaseCallback_RenderResult(ByVal Result As App_UseCaseResult)
    If Result.TypeCode = Success Then
        This.Presenter.ShowSuccess Result.Message
    ElseIf Result.TypeCode = BusinessError Then
        This.Presenter.NotifyBusinessError Result.Message
    ElseIf Result.TypeCode = SystemError Then
        This.Presenter.NotifySystemError
    End If
End Sub
