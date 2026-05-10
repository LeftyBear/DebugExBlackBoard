VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
    Counter As App_TimeTableSummaryCounter
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Factory = New Dom_ValueObjectFactory
    Set This.Counter = New App_TimeTableSummaryCounter
End Sub
'
'Public Sub Initialize(ByVal Counter As App_TimeTableSummaryCounter)
'    Set This.Counter = Counter
'    Set This.Factory = Factory
'End Sub

Public Function Format(ByVal TypeCode As Dom_ClassHourTypePolicy, ByVal Aggregate As Dom_DailyClassHour, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim TimeTable() As String
    TimeTable = ToTimeTable(TypeCode, Aggregate, Structure)
    Dim Total() As Long
    ReDim Total(Structure.UpperValue.Item.MainStreamGradeCount)
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
                OutputTable(R, C) = Structure.Subject.FindByIndex(R - 1).Name
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

Private Function ToTimeTable(ByVal TypeCode As Dom_ClassHourTypePolicy, ByVal Aggregate As Dom_DailyClassHour, ByVal Structure As Dom_SchoolStructure) As String()
    Dim MainStreamGradeCount As Long
    MainStreamGradeCount = Structure.UpperValue.Item.MainStreamGradeCount
    Dim PeriodCount As Long
    PeriodCount = Structure.UpperValue.Item.PeriodCount
    Dim Table() As String
    ReDim Table(1 To PeriodCount, 1 To MainStreamGradeCount)
    If Aggregate Is Nothing Then GoTo Finalize
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.Factory.CreateClassHourType(TypeCode)
    Dim G As Long
    For G = 1 To MainStreamGradeCount
        Dim Grade As Dom_Grade
        Set Grade = This.Factory.CreateGrade(G)
        Dim P As Long
        For P = 1 To PeriodCount
            Dim Period As Dom_Period
            Set Period = This.Factory.CreatePeriod(P)
            Dim Condition As App_TimeTableFilterCondition
            Set Condition = This.Factory.CreateTimeTableFilterCondition(ClassHourType, Period, Grade)
            Table(P, G) = This.Counter.CountHours(Aggregate, Condition)
        Next
    Next
Finalize:
    ToTimeTable = Table
End Function
