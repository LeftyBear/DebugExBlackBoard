VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Private Type Member
    ClassHourFactory    As Inf_ClassHourFactory
    ClassHourIDMapper   As Inf_ClassHourIDMapper
    TimeTableFactory    As Inf_TimeTableFactory
    TimeTableIDMapper   As Inf_TimeTableIDMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ClassHourFactory = New Inf_ClassHourFactory
    Set This.ClassHourIDMapper = New Inf_ClassHourIDMapper
    This.ClassHourFactory.Inject This.ClassHourIDMapper
    Set This.TimeTableFactory = New Inf_TimeTableFactory
    Set This.TimeTableIDMapper = New Inf_TimeTableIDMapper
    This.TimeTableFactory.Inject This.TimeTableIDMapper
End Sub

Public Function Create(ByVal DailyRows As Inf_ClassHourRows) As Dom_DailyClassHours
    Dim Result As Dom_DailyClassHours
    Set Result = New Dom_DailyClassHours
    Dim i As Long
    For i = 1 To DailyRows.Count
        If Not This.ClassHourIDMapper.ToDomain(DailyRows.Item(i).RawID) Is Nothing Then
            Dim ClassHour As Dom_ClassHour
            Set ClassHour = This.ClassHourFactory.Create(DailyRows.Item(i))
            If Not ClassHour Is Nothing Then Result.AddClassHour ClassHour
        ElseIf Not This.TimeTableIDMapper.ToDomain(DailyRows.Item(i).RawID) Is Nothing Then
            Dim TimeTable As Dom_TimeTable
            Set TimeTable = This.TimeTableFactory.Create(DailyRows.Item(i))
            If Not TimeTable Is Nothing Then Result.AddTimeTable TimeTable
        End If
    Next
    Set Create = Result
End Function
