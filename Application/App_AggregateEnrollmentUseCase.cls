VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateEnrollmentUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Calculator As Dom_SchoolYearCalculator
    Repository As Dom_IEnrollmentRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Calculator As Dom_SchoolYearCalculator, ByVal Repository As Dom_IEnrollmentRepository)
    Set This.Calculator = Calculator
    Set This.Repository = Repository
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_EnrollmentSummary
    Dim SchoolYear As Long
    SchoolYear = This.Calculator.Calculate(TargetDate)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SchoolYear)
    Dim PackColumnDTO As App_PackColumnDTOUseCase
    Set PackColumnDTO = New App_PackColumnDTOUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = PackColumnDTO.Execute(CSVDTO)
    Dim GenerateAggregate As App_GenerateEntityAggUseCase
    Set GenerateAggregate = New App_GenerateEntityAggUseCase
    GenerateAggregate.Initialize New Dom_EnrollmentColumnFactory
    Dim Aggregate As Dom_EntityAggregate
    Set Aggregate = GenerateAggregate.Execute(ColumnDTOContainer)
    Dim SummaryService As Dom_EnrollmentSummaryService
    Set SummaryService = New Dom_EnrollmentSummaryService
    Set Execute = SummaryService.Execute(Aggregate)
'    Dim Records As VBA.Collection
'    Set Records = Summary.ToRecordCollection
'    Dim Formatter As App_EnrollmentSummaryFormatter
'    Set Formatter = New App_EnrollmentSummaryFormatter
End Function
