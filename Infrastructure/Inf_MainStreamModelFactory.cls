VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Row As Inf_MainStreamRow) As App_MainStreamReadModel
    Dim Result As App_MainStreamReadModel
    Set Result = New App_MainStreamReadModel
    Result.UpperGrade = Row.UpperGrade
    Result.UpperClassNo = Row.UpperClassNo
    Set Create = Result
End Function
