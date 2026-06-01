Attribute VB_Name = "Inf_SchoolEventModelFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Rows As Inf_SchoolEventRows) As App_SchoolEventReadModels
    Dim Result As App_SchoolEventReadModels
    Set Result = New App_SchoolEventReadModels
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SchoolEventRow
        Set Row = Rows.Item(i)
        Dim Model As App_SchoolEventReadModel
        Set Model = New App_SchoolEventReadModel
        Model.Name = Row.Name
        Result.Add Model
    Next
    Set Create = Result
End Function
