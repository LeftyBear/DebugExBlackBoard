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
    Factory         As Dom_ValueObjectFactory
    TypeMapper      As Inf_ClassHourTypeMapper
    PeriodMapper    As Inf_PeriodMapper
    GradeMapper     As Inf_GradeMapper
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
    Set This.TypeMapper = New Inf_ClassHourTypeMapper
    Set This.PeriodMapper = New Inf_PeriodMapper
    Set This.GradeMapper = New Inf_GradeMapper
    This.TypeMapper.Initialize This.Factory
    This.GradeMapper.Initialize This.Factory
    This.PeriodMapper.Initialize This.Factory
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
    Dim Parts() As String
    Parts = VBA.Split(Raw, DELIMITER)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        Dim ClassHourType As Dom_ClassHourType
        Dim Grade As Dom_Grade
        Dim Period As Dom_Period
        If ClassHourType Is Nothing Then
            Set ClassHourType = This.TypeMapper.ToDomain(Part)
        ElseIf Grade Is Nothing Then
            Set Grade = This.GradeMapper.ToDomain(Part)
        ElseIf Period Is Nothing Then
            Set Period = This.PeriodMapper.ToDomain(Part)
        End If
    Next
    Set ToDomain = This.Factory.CreateTimeTableID(ClassHourType, Period, Grade)
End Function
