VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentAggregateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Builder")
Option Explicit

Public Function Build(ByVal RowsOfColumnDtos As VBA.Collection) As Dom_EnrollmentAggregate
    If RowsOfColumnDtos Is Nothing Then Err.Raise 5, , "RowsOfColumnDtos is Nothing."
    If RowsOfColumnDtos.Count = 0 Then Err.Raise 5, , "RowsOfColumnDtos is empty."
    Dim EntityRows As VBA.Collection
    Set EntityRows = New VBA.Collection
    Dim i As Long
    For i = 1 To RowsOfColumnDtos.Count
        EntityRows.Add BuildRow(RowsOfColumnDtos.Item(i))
    Next
    Dim Aggregate As Dom_EnrollmentAggregate
    Set Aggregate = New Dom_EnrollmentAggregate
    Aggregate.Initialize EntityRows
    Set Build = Aggregate
End Function

Private Function BuildRow(ByVal ColumnDTOs As VBA.Collection) As Dom_EnrollmentRow
    If ColumnDTOs Is Nothing Then Err.Raise 5, , "ColumnDtos is Nothing."
    Dim EntityColumns As VBA.Collection
    Set EntityColumns = New VBA.Collection
    Dim Factory As Dom_EnrollmentColumnFactory
    Set Factory = New Dom_EnrollmentColumnFactory
    Dim i As Long
    For i = 1 To ColumnDTOs.Count
        EntityColumns.Add Factory.Create(ColumnDTOs.Item(i))
    Next
    Dim EntityRow As Dom_EnrollmentRow
    Set EntityRow = New Dom_EnrollmentRow
    EntityRow.Initialize EntityColumns
    Set BuildRow = EntityRow
End Function
