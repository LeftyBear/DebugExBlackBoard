VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_TimeTableIDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Type Member
    TypeMapper      As Inf_ClassHourTypeMapper
    PeriodMapper    As Inf_PeriodMapper
    GradeMapper     As Inf_GradeMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.TypeMapper = New Inf_ClassHourTypeMapper
    Set This.PeriodMapper = New Inf_PeriodMapper
    Set This.GradeMapper = New Inf_GradeMapper
End Sub

Public Function ToRaw(ByVal ID As Dom_TimeTableID) As String
    Dim TypeCode As String
    TypeCode = This.TypeMapper.ToRaw(ID.ClassHourType)
    Dim Period As String
    Period = This.PeriodMapper.ToRaw(ID.Period)
    Dim Grade As String
    Grade = This.GradeMapper.ToRaw(ID.Grade)
    ToRaw = TypeCode & DELIMITER & Period & DELIMITER & Grade
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_TimeTableID
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.TypeMapper.ToDomain(Raw)
    Dim Period As Dom_Period
    Set Period = This.PeriodMapper.ToDomain(Raw)
    Dim Grade As Dom_Grade
    Set Grade = This.GradeMapper.ToDomain(Raw)
    Set ToDomain = Dom_ValueObjectFactory.CreateTimeTableID(ClassHourType, Period, Grade)
End Function
