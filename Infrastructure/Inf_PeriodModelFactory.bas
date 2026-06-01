Attribute VB_Name = "Inf_PeriodModelFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Row As Inf_PeriodRow) As App_PeriodReadModel
    Dim Result As App_PeriodReadModel
    Set Result = New App_PeriodReadModel
    Result.Value = Row.Value
    Set Create = Result
End Function
