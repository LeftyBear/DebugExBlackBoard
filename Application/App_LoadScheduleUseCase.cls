VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadScheduleUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit
Implements App_IUseCaseCallback
Private Type Member
    QueryService    As App_IScheduleQueryService
    Base            As App_BaseUseCase
    Presenter       As App_IDailySchedulePresenter
    TargetDate      As Date
    ViewModels      As Pre_ScheduleViewModels
End Type

Private This As Member

Friend Sub Inject(ByVal QueryService As Inf_ScheduleQueryService, ByVal Base As App_BaseUseCase, ByVal Presenter As App_IDailySchedulePresenter)
    Set This.QueryService = QueryService
    Set This.Base = Base
    Set This.Presenter = Presenter
End Sub

Public Sub SetDate(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
End Sub

Public Property Get ViewModel() As Pre_ScheduleViewModels
    Set ViewModel = This.ViewModels
End Property

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_LoadScheduleUseCase", "TargetDate is not set."
    Dim DailyScheduleRMs As App_ScheduleReadModels
    Set DailyScheduleRMs = This.QueryService.GetByDate(This.TargetDate)
    Set This.ViewModels = This.Presenter.CreateViewModel(DailyScheduleRMs)
End Sub
