VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_GenerateEntityAggUseCase"
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

Public Function Execute(ByVal ColumnDTOContainer As VBA.Collection) As Dom_EntityAggregate
    If ColumnDTOContainer Is Nothing Then Err.Raise AppErrNotingObject, , "ColumnDTOContainer is nothing"
    If ColumnDTOContainer.Count = 0 Then Err.Raise AppErrEmptyObject, , "ColumnDTOContainer is empty"
    Dim ColumnContextContainer As VBA.Collection
    Set ColumnContextContainer = New VBA.Collection
    Dim i As Long
    For i = 1 To ColumnDTOContainer.Count
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(i)
        ColumnContextContainer.Add ColumnContext(ColumnDTOs)
    Next
    Dim Result As Dom_EntityAggregate
    Set Result = New Dom_EntityAggregate
    Result.Initialize ColumnContextContainer
    Set Execute = Result
End Function

Private Function ColumnContext(ByVal ColumnDTOs As VBA.Collection) As Dom_ColumnContext
    Dim EntityColumns As VBA.Collection
    Set EntityColumns = New VBA.Collection
    Dim i As Long
    For i = 1 To ColumnDTOs.Count
        Dim DTO As App_ColumnDTO
        Set DTO = ColumnDTOs.Item(i)
        EntityColumns.Add This.Factory.Create(DTO)
    Next
    Dim Result As Dom_ColumnContext
    Set Result = New Dom_ColumnContext
    Result.Initialize EntityColumns
    Set ColumnContext = Result
End Function
