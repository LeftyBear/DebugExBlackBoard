VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateSubjectUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Repository As Dom_ISubjectRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_ISubjectRepository)
    Set This.Repository = Repository
End Sub

Public Function Execute() As Dom_EntityAggregate
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(Subject)
    Dim PackColumnDTO As App_PackColumnDTOUseCase
    Set PackColumnDTO = New App_PackColumnDTOUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = PackColumnDTO.Execute(CSVDTO)
    Dim GenerateAggregate As App_GenerateEntityAggUseCase
    Set GenerateAggregate = New App_GenerateEntityAggUseCase
    GenerateAggregate.Initialize New Dom_ConfigColumnFactory
    Set Execute = GenerateAggregate.Execute(ColumnDTOContainer)
End Function
