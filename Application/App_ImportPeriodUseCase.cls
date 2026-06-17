VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ImportPeriodUseCase"
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
    ViewModel       As Pre_PeriodViewModel
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

Public Property Get ViewModel() As Pre_PeriodViewModel
    Set ViewModel = This.ViewModel
End Property

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_ImportPeriodUseCase", "TargetDate is not set."
    Dim ReadModels As App_ClassHourReadModels
    Set ReadModels = This.ClassHourQS.GetByDate(This.TargetDate)
    Dim MainStream As App_MainStreamReadModel
    Set MainStream = This.MainStreamQS.GetAll
    Dim Calculator As App_PeriodCalculator
    Set Calculator = New App_PeriodCalculator
    Dim Periods As App_Periods
    Set Periods = Calculator.CountPeriodOfGrade(ReadModels, ExeKeyCode)
    Set This.ViewModel = This.Presenter.CreateViewModel(Periods, MainStream.UpperGrade)
End Sub
