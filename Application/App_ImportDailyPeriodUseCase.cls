VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ImportDailyPeriodUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit
Implements App_IUseCaseCallback
Private Type Member
    ClassHourQS     As App_IClassHourQueryService
    MainStreamQS    As App_IMainStreamQueryService
    Base            As App_BaseUseCase
    Presenter       As App_IDailyPeriodPresenter
    TargetDate      As Date
    ViewModel()     As Variant
End Type

Private This As Member

Friend Sub Inject(ByVal ClassHourQS As App_IClassHourQueryService, ByVal MainStreamQS As App_IMainStreamQueryService, ByVal Base As App_BaseUseCase, ByVal Presenter As App_IDailyPeriodPresenter)
    Set This.ClassHourQS = ClassHourQS
    Set This.MainStreamQS = MainStreamQS
    Set This.Base = Base
    Set This.Presenter = Presenter
End Sub

Public Sub SetDate(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
End Sub

Public Property Get ViewModel() As Variant()
    ViewModel = This.ViewModel
End Property

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_ImportDailyPeriodUseCase", "TargetDate is not set."
    Dim ClassHourReadModels As App_ClassHourReadModels
    Set ClassHourReadModels = This.ClassHourQS.GetByDate(This.TargetDate)
    Dim MainStreamReadModel As App_MainStreamReadModel
    Set MainStreamReadModel = This.MainStreamQS.GetModel
    This.ViewModel = This.Presenter.CreateViewModel(ClassHourReadModels, MainStreamReadModel)
End Sub
