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
    Dim Resolver As App_ClassHourTypeResolver
    Set Resolver = New App_ClassHourTypeResolver
    AddClassHourModel Entries, Resolver, Result
    AddTimTableModel Entries, Resolver, Result
    Set Create = Result
End Function

Private Sub AddClassHourModel(ByVal Entries As Dom_DailyClassHours, ByVal Resolver As App_ClassHourTypeResolver, ByVal ReadModels As App_ClassHourReadModels)
    Dim i As Long
    For i = 1 To Entries.ClassHoursCount
        Dim Model As App_ClassHourReadModel
        Set Model = New App_ClassHourReadModel
        Dim Entry As Dom_ClassHour
        Set Entry = Entries.ClassHoursItem(i)
        Model.TypeKey = Resolver.ParseKey(Entry.ID.ClassHourType.Value)
        Model.Subject = Entry.ID.Subject.Value
        Model.Grade = Entry.ID.Grade.Value
        Model.Value = Entry.Value.Value
        If Entry.ID.ClassHourType.Value = PlanCode Then
            ReadModels.AddClassHourPlan Model
        ElseIf Entry.ID.ClassHourType.Value = ExeCode Then
            ReadModels.AddClassHourExe Model
        End If
    Next
End Sub

Private Sub AddTimTableModel(ByVal Entries As Dom_DailyClassHours, ByVal Resolver As App_ClassHourTypeResolver, ByVal ReadModels As App_ClassHourReadModels)
    Dim i As Long
    For i = 1 To Entries.TimeTablesCount
        Dim Model As App_TimeTableReadModel
        Set Model = New App_TimeTableReadModel
        Dim Entry As Dom_TimeTable
        Set Entry = Entries.TimeTablesItem(i)
        Model.TypeKey = Resolver.ParseKey(Entry.ID.ClassHourType.Value)
        Model.Period = Entry.ID.Period.Value
        Model.Grade = Entry.ID.Grade.Value
        Model.Mark = Entry.Value.Value
        If Entry.ID.ClassHourType.Value = PlanCode Then
            ReadModels.AddTimeTablePlan Model
        ElseIf Entry.ID.ClassHourType.Value = ExeCode Then
            ReadModels.AddTimeTableExe Model
        End If
    Next
End Sub
