VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TotalDailyPeriodUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    ClassHourQS     As App_IClassHourQueryService
    MainStreamQS    As App_IMainStreamQueryService
End Type

Private This As Member

Public Sub Inject(ByVal ClassHourQS As App_IClassHourQueryService, ByVal MainStreamQS As App_IMainStreamQueryService)
    Set This.ClassHourQS = ClassHourQS
    Set This.MainStreamQS = MainStreamQS
End Sub

Public Function Execute(ByVal TargetDate As Date) As App_DailyPeriodReadModels
    Dim ClassHours As App_ClassHourReadModels
    Set ClassHours = This.ClassHourQS.GetByDate(TargetDate)
    Dim MainStream As App_MainStreamReadModel
    Set MainStream = This.MainStreamQS.GetModel
    Dim Calculator As New App_DailyPeriodTotalCalclator
    Dim Mapper As New App_IClassHourTypeMapper
    Set Execute = Calculator.Calclate(ClassHours, Mapper.ExecutionKey, MainStream)
End Function
