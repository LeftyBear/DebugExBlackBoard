VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

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
