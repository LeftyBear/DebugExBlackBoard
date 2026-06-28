VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadClassHourUseCase"
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

Private Sub SetClassHourViewModel(ByVal ClassHourRMs As App_ClassHourReadModels, ByVal GradeCount As Long)
    Dim SubjectRMs As App_SubjectReadModels
    Set SubjectRMs = This.SubjectQS.GetAll
    Dim Calculator As App_ClassHourCalculator
    Set Calculator = New App_ClassHourCalculator
    Dim PlanTotals As App_ClassHourTotals
    Set PlanTotals = Calculator.CalculatePlanTotal(ClassHourRMs, SubjectRMs)
    Dim ExeTotals As App_ClassHourTotals
    Set ExeTotals = Calculator.CalculateExeTotal(ClassHourRMs, SubjectRMs)
    Set This.ClassHourPlanVM = This.ClassHourPre.CreatePlanViewModel(ClassHourRMs, SubjectRMs, PlanTotals, GradeCount)
    Set This.ClassHourExeVM = This.ClassHourPre.CreateExeViewModel(ClassHourRMs, SubjectRMs, ExeTotals, GradeCount)
End Sub

Private Sub SetTimeTableViewModel(ByVal ClassHourRMs As App_ClassHourReadModels, ByVal GradeCount As Long)
    Dim Calculator As App_PeriodCalculator
    Set Calculator = New App_PeriodCalculator
    Dim PlanTotals As App_PeriodTotals
    Set PlanTotals = Calculator.CountPeriodPlan(ClassHourRMs)
    Dim ExeTotals As App_PeriodTotals
    Set ExeTotals = Calculator.CountPeriodExe(ClassHourRMs)
    Dim PeriodRM As App_PeriodReadModel
    Set PeriodRM = This.PeriodQS.GetAll
    Dim PeriodCount As Long
    PeriodCount = PeriodRM.Value
    Set This.TimeTablePlanVM = This.TimeTablePre.CreatePlanViewModel(ClassHourRMs, PlanTotals, GradeCount, PeriodCount)
    Set This.TimeTableExeVM = This.TimeTablePre.CreateExeViewModel(ClassHourRMs, ExeTotals, GradeCount, PeriodCount)
End Sub

Private Sub App_IUseCaseCallback_ExecuteCore()
    If This.TargetDate = 0 Then Err.Raise AppErrNotSetVariable, "App_LoadClassHourUseCase", "TargetDate is not set."
    Dim DailyClassHourRMs As App_ClassHourReadModels
    Set DailyClassHourRMs = This.ClassHourQS.GetByDate(This.TargetDate)
    Dim MainStreamRM As App_MainStreamReadModel
    Set MainStreamRM = This.MainStreamQS.GetAll
    Dim GradeCount As Long
    GradeCount = MainStreamRM.UpperGrade
    SetClassHourViewModel DailyClassHourRMs, GradeCount
    SetTimeTableViewModel DailyClassHourRMs, GradeCount
End Sub
