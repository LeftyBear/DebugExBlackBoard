VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ImportClassHourUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit
Implements App_IUseCaseCallback
Private Type Member
    ClassHourQS     As App_IClassHourQueryService
    SubjectQS       As App_ISubjectQueryService
    PeriodQS        As App_IPeriodQueryService
    MainStreamQS    As App_IMainStreamQueryService
    Base            As App_BaseUseCase
    ClassHourPre    As App_IDailyClassHourPresenter
    TimeTablePre    As App_IDailyTimeTablePresenter
    TargetDate      As Date
    ClassHourPlanVM As Pre_TableViewModel
    ClassHourExeVM  As Pre_TableViewModel
    TimeTablePlanVM As Pre_TableViewModel
    TimeTableExeVM  As Pre_TableViewModel
End Type

Private This As Member

Friend Sub Inject(ByVal ClassHourQS As App_IClassHourQueryService, ByVal SubjectQS As App_ISubjectQueryService, ByVal PeriodQS As App_IPeriodQueryService, ByVal MainStreamQS As App_IMainStreamQueryService, ByVal Base As App_BaseUseCase, ByVal ClassHourPre As App_IDailyClassHourPresenter, ByVal TimeTablePre As App_IDailyTimeTablePresenter)
    Set This.ClassHourQS = ClassHourQS
    Set This.SubjectQS = SubjectQS
    Set This.PeriodQS = PeriodQS
    Set This.MainStreamQS = MainStreamQS
    Set This.Base = Base
    Set This.ClassHourPre = ClassHourPre
    Set This.TimeTablePre = TimeTablePre
End Sub

Public Sub SetDate(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
End Sub

Public Property Get ClassHourPlanVM() As Pre_TableViewModel
    Set ClassHourPlanVM = This.ClassHourPlanVM
End Property

Public Property Get ClassHourExeVM() As Pre_TableViewModel
    Set ClassHourExeVM = This.ClassHourExeVM
End Property

Public Property Get TimeTablePlanVM() As Pre_TableViewModel
    Set TimeTablePlanVM = This.TimeTablePlanVM
End Property

Public Property Get TimeTableExeVM() As Pre_TableViewModel
    Set TimeTableExeVM = This.TimeTableExeVM
End Property

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_ImportClassHourUseCase", "TargetDate is not set."
    Dim ReadModels As App_ClassHourReadModels
    Set ReadModels = This.ClassHourQS.GetByDate(This.TargetDate)
    Dim Subjects As App_SubjectReadModels
    Set Subjects = This.SubjectQS.GetAll
    Dim Resolver As App_ClassHourTypeResolver
    Set Resolver = New App_ClassHourTypeResolver
    Dim PlanKey As String
    PlanKey = Resolver.ParseKey(PlanCode)
    Dim ClassHourCalc As App_ClassHourCalculator
    Set ClassHourCalc = New App_ClassHourCalculator
    Dim MainStream As App_MainStreamReadModel
    Set MainStream = This.MainStreamQS.GetAll
    Set This.ClassHourPlanVM = CreateClassHourViewModel(ReadModels, Subjects, PlanKey, ClassHourCalc, MainStream.UpperGrade)
    Dim ExeKey As String
    ExeKey = Resolver.ParseKey(ExeCode)
    Set This.ClassHourExeVM = CreateClassHourViewModel(ReadModels, Subjects, ExeKey, ClassHourCalc, MainStream.UpperGrade)
    Dim Period As App_PeriodReadModel
    Set Period = This.PeriodQS.GetAll
    Dim PeriodCalc As App_PeriodCalculator
    Set PeriodCalc = New App_PeriodCalculator
    Set This.TimeTablePlanVM = CreateTimeTableViewModel(ReadModels, PlanKey, PeriodCalc, Period.Value, MainStream.UpperGrade)
    Set This.TimeTableExeVM = CreateTimeTableViewModel(ReadModels, ExeKey, PeriodCalc, Period.Value, MainStream.UpperGrade)
End Sub

Private Function CreateClassHourViewModel(ByVal Models As App_ClassHourReadModels, ByVal Subjects As App_SubjectReadModels, ByVal TypeKey As String, ByVal Calculator As App_ClassHourCalculator, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Totals As App_ClassHourTotals
    Set Totals = Calculator.CalculateTotal(Models, Subjects, TypeKey)
    Set CreateClassHourViewModel = This.ClassHourPre.CreateViewModel(Models, Subjects, Totals, GradeCount)
End Function

Private Function CreateTimeTableViewModel(ByVal Models As App_ClassHourReadModels, ByVal TypeKey As String, ByVal Calculator As App_PeriodCalculator, ByVal PeriodCount As Long, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Totals As App_PeriodTotals
    Set Totals = Calculator.CountPeriod(Models, TypeKey)
    Set CreateTimeTableViewModel = This.TimeTablePre.CreateViewModel(Models, Totals, PeriodCount, GradeCount)
End Function
