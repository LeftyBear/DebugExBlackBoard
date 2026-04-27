VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Factory As App_ValueObjectFactory
End Type

Private This As Member

Private Sub Initialize(ByVal Factory As App_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByVal NormDate As Date, ByVal Row As App_TimeTableRow) As Dom_TimeTable
    Dim ID As Dom_TimeTableID
    Set ID = CreateID(Row)
    If ID.IsNoMeaning Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = This.Factory.CreateDateID(NormDate)
    Dim Value As Dom_TimeTableValue
    Set Value = This.Factory.CreateTimeTableValue(Row.Value)
    Dim Entity As Dom_TimeTable
    Set Entity = New Dom_TimeTable
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function

Private Function CreateID(ByVal Row As App_TimeTableRow) As Dom_TimeTableID
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.Factory.CreateClassHourType(Row.TypeCode)
    Dim Period As Dom_Period
    Set Period = This.Factory.CreatePeriod(Row.Period)
    Dim Grade As Dom_Grade
    Set Grade = This.Factory.CreateGrade(Row.Grade)
    Set CreateID = This.Factory.CreateTimeTableID(ClassHourType, Period, Grade)
End Function
