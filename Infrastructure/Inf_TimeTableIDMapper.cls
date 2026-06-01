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
    Set ToDomain = Dom_TimeTableIDFactory.Create(ClassHourType, Period, Grade)
End Function
