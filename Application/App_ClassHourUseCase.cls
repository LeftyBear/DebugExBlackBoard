VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
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

Public Function Excute(ByVal TargetDate As Date) As Dom_EntityAggregate
    Dim SchoolYear As Long
    SchoolYear = This.Calculator.Calculate(TargetDate)
    Dim RawRows() As String
    RawRows = This.Repository.Load(SchoolYear)
    Dim RowOfColumnDTOs As VBA.Collection
    Set RowOfColumnDTOs = BuildRows(RawRows)
    Dim AggregateBuilder As App_EntityAggregateBuilder
    Set AggregateBuilder = New App_EntityAggregateBuilder
    AggregateBuilder.Initialize New Dom_ClassHourColumnFactory
    Set Excute = AggregateBuilder.Build(RowOfColumnDTOs)
End Function

Private Function BuildRows(ByRef RawRows() As String) As VBA.Collection
    Dim Header() As String
    ReDim Header(LBound(RawRows, 2) To UBound(RawRows, 2))
    Dim C As Long
    For C = LBound(RawRows, 2) To UBound(RawRows, 2)
        Header(C) = RawRows(LBound(RawRows), C)
    Next
    Dim RowOfColumnDTOs As VBA.Collection
    Set RowOfColumnDTOs = New VBA.Collection
    Dim R As Long
    For R = LBound(RawRows) To UBound(RawRows)
        For C = LBound(RawRows, 2) To UBound(RawRows, 2)
            Dim DTO As App_ColumnDTO
            Set DTO = New App_ColumnDTO
            DTO.Initialize Header(C), RawRows(R, C)
            RowOfColumnDTOs.Add DTO
        Next
    Next
    Set BuildRows = RowOfColumnDTOs
End Function

