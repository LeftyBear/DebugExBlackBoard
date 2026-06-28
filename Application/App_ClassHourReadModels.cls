VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourReadModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTOCollection")
Option Explicit

Private Type Member
    ClassHourPlanModels As VBA.Collection
    ClassHourExeModels  As VBA.Collection
    TimeTablePlanModels As VBA.Collection
    TimeTableExeModels  As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ClassHourPlanModels = New VBA.Collection
    Set This.ClassHourExeModels = New VBA.Collection
    Set This.TimeTablePlanModels = New VBA.Collection
    Set This.TimeTableExeModels = New VBA.Collection
End Sub

Public Sub AddClassHourPlan(ByVal Model As App_ClassHourReadModel)
    This.ClassHourPlanModels.Add Model
End Sub

Public Sub AddClassHourExe(ByVal Model As App_ClassHourReadModel)
    This.ClassHourExeModels.Add Model
End Sub

Public Sub AddTimeTablePlan(ByVal Model As App_TimeTableReadModel)
    This.TimeTablePlanModels.Add Model
End Sub

Public Sub AddTimeTableExe(ByVal Model As App_TimeTableReadModel)
    This.TimeTableExeModels.Add Model
End Sub

Public Function ClassHourPlanCount() As Long
    ClassHourPlanCount = This.ClassHourPlanModels.Count
End Function

Public Function ClassHourExeCount() As Long
    ClassHourExeCount = This.ClassHourExeModels.Count
End Function

Public Function TimeTablePlanCount() As Long
    TimeTablePlanCount = This.TimeTablePlanModels.Count
End Function

Public Function TimeTableExeCount() As Long
    TimeTableExeCount = This.TimeTableExeModels.Count
End Function

Public Function GetClassHourPlan(ByVal INDEX As Long) As App_ClassHourReadModel
    Set GetClassHourPlan = This.ClassHourPlanModels.Item(INDEX)
End Function

Public Function GetClassHourExe(ByVal INDEX As Long) As App_ClassHourReadModel
    Set GetClassHourExe = This.ClassHourExeModels.Item(INDEX)
End Function

Public Function GetTimeTablePlan(ByVal INDEX As Long) As App_TimeTableReadModel
    Set GetTimeTablePlan = This.TimeTablePlanModels.Item(INDEX)
End Function

Public Function GetTimeTableExe(ByVal INDEX As Long) As App_TimeTableReadModel
    Set GetTimeTableExe = This.TimeTableExeModels.Item(INDEX)
End Function
