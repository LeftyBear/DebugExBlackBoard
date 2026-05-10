VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DailyTimeTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Private Type Member
    ValueObjectFactory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal ValueObjectFactory As Dom_ValueObjectFactory)
    Set This.ValueObjectFactory = ValueObjectFactory
End Sub

Public Function Create(ByVal Row As App_ClassHourRow) As Dom_TimeTable
    Dim DateID As Dom_DateID
    Set DateID = This.ValueObjectFactory.CreateDateID(Row.Value)
    Dim ID As Dom_TimeTableID
    Set ID = CreateID(Row)
    If ID Is Nothing Then Exit Function
    Dim Value As Dom_TimeTableValue
    Set Value = This.ValueObjectFactory.CreateTimeTableValue(Row.Value)
    Dim Entity As Dom_TimeTable
    Set Entity = New Dom_TimeTable
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function

Private Function CreateID(ByVal Row As App_ClassHourRow) As Dom_TimeTableID
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.ValueObjectFactory.CreateClassHourType(Row.TypeCode)
    If ClassHourType.Value = 0 Then Exit Function
    Dim Period As Dom_Period
    Set Period = This.ValueObjectFactory.CreatePeriod(Row.Period)
    If Period.Value = 0 Then Exit Function
    Dim Grade As Dom_Grade
    Set Grade = This.ValueObjectFactory.CreateGrade(Row.Grade)
    If Grade.Value = 0 Then Exit Function
    Set CreateID = This.ValueObjectFactory.CreateTimeTableID(ClassHourType, Period, Grade)
End Function
