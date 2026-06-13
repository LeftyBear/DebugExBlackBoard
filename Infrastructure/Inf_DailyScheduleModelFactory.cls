VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyScheduleModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Entries As Dom_DailySchedules) As App_ScheduleReadModels
    Dim Result As App_ScheduleReadModels
    Set Result = New App_ScheduleReadModels
    Dim i As Long
    For i = 1 To Entries.Count
        Dim Model As App_ScheduleReadModel
        Set Model = New App_ScheduleReadModel
        Dim Entry As Dom_Schedule
        Set Entry = Entries.Item(i)
        Model.Name = Entry.Name.Value
        Model.Value = Entry.Value.Value
        Result.Add Model
    Next
    Set Create = Result
End Function
