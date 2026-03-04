VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EntityAggregateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Builder")
Option Explicit

Private Type Member
    Factory As Dom_IEntityColumnFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_IEntityColumnFactory)
    Set This.Factory = Factory
End Sub

Public Function Build(ByVal RowOfColumnDTOs As VBA.Collection) As Dom_EntityAggregate
    If RowOfColumnDTOs Is Nothing Then Err.Raise AppErrNotingObject, , "RowOfColumnDTOs is nothing"
    If RowOfColumnDTOs.Count = 0 Then Err.Raise AppErrEmptyObject, , "RowOfColumnDTOs is empty"
    Dim EntityRows As VBA.Collection
    Set EntityRows = New VBA.Collection
    Dim i As Long
    For i = 1 To RowOfColumnDTOs.Count
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = RowOfColumnDTOs.Item(i)
        EntityRows.Add BuildRow(ColumnDTOs)
    Next
    Dim Aggregate As Dom_EntityAggregate
    Set Aggregate = New Dom_EntityAggregate
    Aggregate.Initialize EntityRows
    Set Build = Aggregate
End Function

Private Function BuildRow(ByVal ColumnDTOs As VBA.Collection) As Dom_EntityRow
    Dim EntityColumns As VBA.Collection
    Set EntityColumns = New VBA.Collection
    Dim i As Long
    For i = 1 To ColumnDTOs.Count
        Dim ColumnDTO As App_ColumnDTO
        Set ColumnDTO = ColumnDTOs.Item(i)
        EntityColumns.Add This.Factory.Create(ColumnDTO)
    Next
    Dim EntityRow As Dom_EntityRow
    Set EntityRow = New Dom_EntityRow
    EntityRow.Initialize EntityColumns
    Set BuildRow = EntityRow
End Function
