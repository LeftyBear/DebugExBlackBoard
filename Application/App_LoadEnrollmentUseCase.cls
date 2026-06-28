VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadEnrollmentUseCase"
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
    ViewModels      As Pre_TableViewModel
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

Public Property Get ViewModel() As Pre_TableViewModel
    Set ViewModel = This.ViewModels
End Property

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_ImportEnrollmentUseCase", "TargetDate is not set."
    Dim DailyEnrollmentRMs As App_EnrollmentReadModels
    Set DailyEnrollmentRMs = This.EnrollmentQS.GetByDate(This.TargetDate)
    Dim StreamTypeReso As App_StreamTypeResolver
    Set StreamTypeReso = New App_StreamTypeResolver
    Dim MainKey As String
    MainKey = StreamTypeReso.ParseKey(MainCode)
    Dim SpecialKey As String
    SpecialKey = StreamTypeReso.ParseKey(SpecialCode)
    Dim GenderReso As App_GenderResolver
    Set GenderReso = New App_GenderResolver
    Dim MaleKey As String
    MaleKey = GenderReso.ParseKey(MaleCode)
    Dim FemaleKey As String
    FemaleKey = GenderReso.ParseKey(FemaleCode)
    Dim Calculator As App_EnrollmentCalculator
    Set Calculator = New App_EnrollmentCalculator
    Dim EnrollmentTotals As App_EnrollmentTotals
    Set EnrollmentTotals = Calculator.Calculate(DailyEnrollmentRMs, MainKey, SpecialKey, MaleKey, FemaleKey)
    Dim MainStreamRM As App_MainStreamReadModel
    Set MainStreamRM = This.MainStreamQS.GetAll
    Dim SpecialStreamRMs As App_SpecialStreamReadModels
    Set SpecialStreamRMs = This.SpecialStreamQS.GetAll
    Set This.ViewModels = This.Presenter.CreateViewModel(DailyEnrollmentRMs, EnrollmentTotals, MainStreamRM, SpecialStreamRMs)
End Sub
