VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyClassHourModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Entries As Dom_DailyClassHours) As App_ClassHourReadModels
    Dim Result As App_ClassHourReadModels
    Set Result = New App_ClassHourReadModels
    Dim Resolver As App_ClassHourTypeModelResolver
    Set Resolver = New App_ClassHourTypeModelResolver
    Dim i As Long
    For i = 1 To Entries.ClassHoursCount
        Dim ClassHourModel As App_ClassHourReadModel
        Set ClassHourModel = New App_ClassHourReadModel
        Dim ClassHour As Dom_ClassHour
        Set ClassHour = Entries.ClassHoursItem(i)
        ClassHourModel.Key = Resolver.ToModel(ClassHour.ID.ClassHourType)
        ClassHourModel.Subject = ClassHour.ID.Subject.Value
        ClassHourModel.Grade = ClassHour.ID.Grade.Value
        ClassHourModel.Value = ClassHour.Value.Value
        Result.AddClassHour ClassHourModel
    Next
    For i = 1 To Entries.TimeTablesCount
        Dim TimeTableModel As App_TimeTableReadModel
        Set TimeTableModel = New App_TimeTableReadModel
        Dim TimeTable As Dom_TimeTable
        Set TimeTable = Entries.TimeTablesItem(i)
        TimeTableModel.Key = Resolver.ToModel(TimeTable.ID.ClassHourType)
        TimeTableModel.Period = TimeTable.ID.Period.Value
        TimeTableModel.Grade = TimeTable.ID.Grade.Value
        TimeTableModel.Mark = TimeTable.Value.Value
        Result.AddTimeTable TimeTableModel
    Next
    Set Create = Result
End Function
