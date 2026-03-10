VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateClassHourUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Calculator As Dom_SchoolYearCalculator
    Repository As Dom_IClassHourRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Calculator As Dom_SchoolYearCalculator, ByVal Repository As Dom_IClassHourRepository)
    Set This.Calculator = Calculator
    Set This.Repository = Repository
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_EntityAggregate
    Dim SchoolYear As Long
    SchoolYear = This.Calculator.Calculate(TargetDate)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SchoolYear)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim GenerateAggregate As App_GenerateEntityAggUseCase
    Set GenerateAggregate = New App_GenerateEntityAggUseCase
    GenerateAggregate.Initialize New Dom_ClassHourColumnFactory
    Dim Aggregate As Dom_EntityAggregate
    Set Aggregate = GenerateAggregate.Execute(ColumnDTOContainer)
End Function
