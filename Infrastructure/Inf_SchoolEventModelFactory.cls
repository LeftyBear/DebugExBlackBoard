VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolEventModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

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
