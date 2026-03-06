VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LimitValueUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Repository As Dom_ILimitValueRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_ILimitValueRepository)
    Set This.Repository = Repository
End Sub

Public Function Excute() As Dom_EntityAggregate
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(LimitValue)
    Dim HeaderParser As App_HeaderParser
    Set HeaderParser = New App_HeaderParser
    Dim RowOfColumnDTOs As VBA.Collection
    Set RowOfColumnDTOs = HeaderParser.Parse(CSVDTO)
    Dim AggregateBuilder As App_EntityAggregateBuilder
    Set AggregateBuilder = New App_EntityAggregateBuilder
    AggregateBuilder.Initialize New Dom_ConfigColumnFactory
    Set Excute = AggregateBuilder.Build(RowOfColumnDTOs)
End Function


