Attribute VB_Name = "Inf_SpecialStreamModelFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Rows As Inf_SpecialStreamRows) As App_SpecialStreamReadModels
    Dim Result As App_SpecialStreamReadModels
    Set Result = New App_SpecialStreamReadModels
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SpecialStreamRow
        Set Row = Rows.Item(i)
        Dim Model As App_SpecialStreamReadModel
        Set Model = New App_SpecialStreamReadModel
        Model.Name = Row.Name
        Result.Add Model
    Next
    Set Create = Result
End Function
