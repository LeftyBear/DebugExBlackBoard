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
    ClassHourModels As VBA.Collection
    TimeTableModels As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ClassHourModels = New VBA.Collection
    Set This.TimeTableModels = New VBA.Collection
End Sub

Public Sub AddClassHour(ByVal Model As App_ClassHourReadModel)
    This.ClassHourModels.Add Model
End Sub

Public Sub AddTimeTable(ByVal Model As App_TimeTableReadModel)
    This.TimeTableModels.Add Model
End Sub

Public Function ClassHourCount() As Long
    ClassHourCount = This.ClassHourModels.Count
End Function

Public Function TimeTableCount() As Long
    TimeTableCount = This.TimeTableModels.Count
End Function

Public Function ClassHourItem(ByVal Index As Long) As App_ClassHourReadModel
    Set ClassHourItem = This.ClassHourModels.Item(Index)
End Function

Public Function TimeTableItem(ByVal Index As Long) As App_TimeTableReadModel
    Set TimeTableItem = This.TimeTableModels.Item(Index)
End Function
