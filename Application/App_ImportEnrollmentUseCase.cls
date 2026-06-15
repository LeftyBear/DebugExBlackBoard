VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ImportEnrollmentUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit
Implements App_IUseCaseCallback
Private Type Member
    EnrollmentQS    As App_IEnrollmentQueryService
    MainStreamQS    As App_IMainStreamQueryService
    SpecialStreamQS As App_ISpecialStreamQueryService
    Base            As App_BaseUseCase
    Presenter       As App_IDailyEnrollmentPresenter
    TargetDate      As Date
    ViewModels      As Pre_EnrollmentViewModel
End Type

Private This As Member

Friend Sub Inject(ByVal EnrollmentQS As Inf_EnrollmentQueryService, ByVal MainStreamQS As App_IMainStreamQueryService, ByVal SpecialStreamQS As App_ISpecialStreamQueryService, ByVal Base As App_BaseUseCase, ByVal Presenter As App_IDailyEnrollmentPresenter)
    Set This.EnrollmentQS = EnrollmentQS
    Set This.MainStreamQS = MainStreamQS
    Set This.SpecialStreamQS = SpecialStreamQS
    Set This.Base = Base
    Set This.Presenter = Presenter
End Sub

Public Sub SetDate(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
End Sub

Public Property Get ViewModel() As Pre_EnrollmentViewModel
    Set ViewModel = This.ViewModels
End Property

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_ImportEnrollmentUseCase", "TargetDate is not set."
    Dim ReadModels As App_EnrollmentReadModels
    Set ReadModels = This.EnrollmentQS.GetByDate(This.TargetDate)
    Dim MainStream As App_MainStreamReadModel
    Set MainStream = This.MainStreamQS.GetAll
    Dim SpecialStreams As App_SpecialStreamReadModels
    Set SpecialStreams = This.SpecialStreamQS.GetAll
    Dim Totals As App_EnrollmentTotals
    Set Totals = New App_EnrollmentTotals
    With New App_EnrollmentCalculator
        .CalculateClassTotal ReadModels, MainKeyCode, Totals
        .CalculateGenderTotal ReadModels, MainKeyCode, Male, Totals
        .CalculateGenderTotal ReadModels, MainKeyCode, Female, Totals
        .CalculateClassTotal ReadModels, SpecialKeyCode, Totals
        .CalculateGenderTotal ReadModels, SpecialKeyCode, Male, Totals
        .CalculateGenderTotal ReadModels, SpecialKeyCode, Female, Totals
    End With
    Set This.ViewModels = This.Presenter.CreateViewModel(ReadModels, Totals, MainStream.UpperGrade, MainStream.UpperClassNo, SpecialStreams.ClassNoCount)
End Sub
