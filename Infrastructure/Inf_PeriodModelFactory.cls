VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_PeriodModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Row As Inf_PeriodRow) As App_PeriodReadModel
    Dim Result As App_PeriodReadModel
    Set Result = New App_PeriodReadModel
    Result.Value = Row.Value
    Set Create = Result
End Function
