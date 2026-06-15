VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_PeriodCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Calculator")
Option Explicit

Public Function CountPeriodOfGrade(ByVal ReadModels As App_ClassHourReadModels, ByVal KeyCode As App_ClassHourTypeKeyCode) As App_Periods
    If ReadModels Is Nothing Then Exit Function
    Dim Result As App_Periods
    Set Result = New App_Periods
    Dim Resolver As App_ClassHourTypeResolver
    Set Resolver = New App_ClassHourTypeResolver
    Dim TypeKey As String
    TypeKey = Resolver.ParseKey(KeyCode)
    Dim FilterdModels As VBA.Collection
    Set FilterdModels = FilterByTypeKey(ReadModels, TypeKey)
    Dim Model As App_TimeTableReadModel
    For Each Model In FilterdModels
        If Model.Mark <> vbNullString Then
            If Result.Item(Model.Grade) = 0 Then
                Result.Add Model.Grade, 1
            Else
                Result.Add Model.Grade, Result.Item(Model.Grade) + 1
            End If
        End If
    Next
    Set CountPeriodOfGrade = Result
End Function

Private Function FilterByTypeKey(ByVal ReadModels As App_ClassHourReadModels, ByVal TypeKey As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To ReadModels.TimeTableCount
        Dim Model As App_TimeTableReadModel
        Set Model = ReadModels.TimeTableItem(i)
        If Model.TypeKey = TypeKey Then
            Result.Add Model
        End If
    Next
    Set FilterByTypeKey = Result
End Function
