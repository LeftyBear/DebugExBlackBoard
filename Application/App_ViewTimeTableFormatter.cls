VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewTimeTableFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Formatter")
Option Explicit

Private Type Member
    Calculator  As Dom_TimeTableCalculator
    Factory     As App_ValueObjectFactory
    Service     As App_TimeTableService
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Calculator = New Dom_TimeTableCalculator
    Set This.Factory = New App_ValueObjectFactory
    Set This.Service = New App_TimeTableService
End Sub

Public Function Format(ByVal ClassHourTypeEnum As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim TimeTable() As String
    TimeTable = ToTimeTable(ClassHourTypeEnum, ClassHour, TargetDate, Structure)
    Dim Total() As Long
    ReDim Total(Structure.MainStreamGradeCount)
    Dim OutputTable() As Variant
    ReDim OutputTable(1 To UBound(TimeTable, 1) + 1, 1 To UBound(TimeTable, 2) + 1)
    Dim C As Long
    For C = 1 To UBound(TimeTable, 2) + 1
        If C = 1 Then
            OutputTable(1, C) = vbNullString
        Else
            OutputTable(1, C) = CStr(C - 1) & jpnNen
        End If
    Next
    Dim R As Long
    For R = 2 To UBound(TimeTable, 1) + 1
        For C = 1 To UBound(TimeTable, 2) + 1
            If R < UBound(TimeTable, 1) + 1 And C = 1 Then
                OutputTable(R, C) = Structure.SubjectNameOf(R - 1)
            ElseIf R < UBound(TimeTable, 1) + 1 Then
                OutputTable(R, C) = TimeTable(R - 1, C - 1)
                If OutputTable(R, C) <> vbNullString Then Total(C - 1) = Total(C - 1) + 1
            ElseIf R = UBound(TimeTable, 1) + 1 And C = 1 Then
                OutputTable(R, C) = jpnGoukei
            ElseIf R = UBound(TimeTable, 1) + 1 Then
                OutputTable(R, C) = Total(C - 1)
            End If
        Next
    Next
    Format = OutputTable
End Function

Private Function ToTimeTable(ByVal ClassHourTypeEnum As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Structure As Dom_SchoolStructure) As String()
    Dim Table() As String
    ReDim Table(1 To Structure.PeriodCount, 1 To Structure.MainStreamGradeCount)
    If ClassHour Is Nothing Then GoTo Finalize
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.Factory.CreateClassHourType(ClassHourTypeEnum)
    Dim G As Long
    For G = 1 To Structure.MainStreamGradeCount
        Dim Grade As Dom_Grade
        Set Grade = This.Factory.CreateGrade(G)
        Dim P As Long
        For P = 1 To Structure.PeriodCount
            Dim Period As Dom_Period
            Set Period = This.Factory.CreatePeriod(P)
            Dim Filter As Dom_TimeTableFilter
            Set Filter = CreateTimeTableFilter(ClassHourType, Period, Grade)
            Table(P, G) = This.Service.GetDailyMarkCount(ClassHour, TargetDate, This.Calculator, Filter)
        Next
    Next
Finalize:
    ToTimeTable = Table
End Function

Private Function CreateTimeTableFilter(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_TimeTableFilter
    With New Dom_TimeTableFilter
        Set CreateTimeTableFilter = .Create(ClassHourType, Period, Grade)
    End With
End Function

