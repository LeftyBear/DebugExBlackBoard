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

Public Function Format(ByVal ClassHourType As Dom_ClassHourType, ByVal Record As Dom_ClassHourRecord, ByVal Config As Dom_SchoolConfig) As Variant()
    Dim TimeTable() As String
    TimeTable = ToTimeTable(ClassHourType, Record, Config)
    Dim Total() As Long
    ReDim Total(Config.MainStreamGradeCount)
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
                OutputTable(R, C) = Config.SubjectName(R - 1)
            ElseIf R < UBound(TimeTable, 1) + 1 Then
                OutputTable(R, C) = TimeTable(R - 1, C - 1)
                If Not OutputTable(R, C) Like vbNullString Then Total(C - 1) = Total(C - 1) + 1
            ElseIf R = UBound(TimeTable, 1) + 1 And C = 1 Then
                OutputTable(R, C) = jpnGoukei
            ElseIf R = UBound(TimeTable, 1) + 1 Then
                OutputTable(R, C) = Total(C - 1)
            End If
        Next
    Next
    Format = OutputTable
End Function

Private Function ToTimeTable(ByVal ClassHourType As Dom_ClassHourType, ByVal Record As Dom_ClassHourRecord, ByVal Config As Dom_SchoolConfig) As String()
    Dim Table() As String
    ReDim Table(1 To Config.PeriodCount, 1 To Config.MainStreamGradeCount)
    Dim Grade As Long
    For Grade = 1 To Config.MainStreamGradeCount
        Dim Period As Long
        For Period = 1 To Config.PeriodCount
            Table(Period, Grade) = Record.TimeTable(ClassHourType, Period, Grade)
        Next
    Next
    ToTimeTable = Table
End Function
