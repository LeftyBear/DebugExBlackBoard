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
    Dim DTO As App_EnrollmentFormatterDTO
    Set DTO = CreateFormatterDTO
    Dim Totals As App_EnrollmentTotals
    Set Totals = GetTotals(ReadModels, DTO)
    Set This.ViewModels = This.Presenter.CreateViewModel(ReadModels, Totals, DTO)
End Sub

Private Function CreateFormatterDTO() As App_EnrollmentFormatterDTO
    Dim Result As App_EnrollmentFormatterDTO
    Set Result = New App_EnrollmentFormatterDTO
    Dim MainStream As App_MainStreamReadModel
    Set MainStream = This.MainStreamQS.GetAll
    Dim SpecialStreams As App_SpecialStreamReadModels
    Set SpecialStreams = This.SpecialStreamQS.GetAll
    Dim StreamTypeResolver As App_StreamTypeResolver
    Set StreamTypeResolver = New App_StreamTypeResolver
    Dim GenderResolver As App_GenderResolver
    Set GenderResolver = New App_GenderResolver
    Result.MainGradeCount = MainStream.UpperGrade
    Result.MainClassNoCount = MainStream.UpperClassNo
    Result.SpecialClassNoCount = SpecialStreams.ClassNoCount
    Result.MainKey = StreamTypeResolver.ParseKey(MainKeyCode)
    Result.SpecialKey = StreamTypeResolver.ParseKey(SpecialKeyCode)
    Result.MaleKey = GenderResolver.ParseKey(MaleCode)
    Result.FemaleKey = GenderResolver.ParseKey(FemaleCode)
    Set CreateFormatterDTO = Result
End Function

Private Function GetTotals(ByVal ReadModels As App_EnrollmentReadModels, ByVal DTO As App_EnrollmentFormatterDTO) As App_EnrollmentTotals
    Dim Result As App_EnrollmentTotals
    Set Result = New App_EnrollmentTotals
    Dim Calculator As App_EnrollmentCalculator
    Set Calculator = New App_EnrollmentCalculator
    Calculator.CalculateClassTotal ReadModels, DTO.MainKey, Result
    Calculator.CalculateGenderTotal ReadModels, DTO.MainKey, DTO.MaleKey, Result
    Calculator.CalculateGenderTotal ReadModels, DTO.MainKey, DTO.FemaleKey, Result
    Calculator.CalculateClassTotal ReadModels, DTO.SpecialKey, Result
    Calculator.CalculateGenderTotal ReadModels, DTO.SpecialKey, DTO.MaleKey, Result
    Calculator.CalculateGenderTotal ReadModels, DTO.SpecialKey, DTO.FemaleKey, Result
End Function
