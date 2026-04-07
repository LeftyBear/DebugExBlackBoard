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

Public Function Format(ByVal ClassHourType As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourAggregate, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim TimeTable() As String
    TimeTable = ToTimeTable(ClassHourType, ClassHour, Structure)
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

Private Function ToTimeTable(ByVal ClassHourType As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourColumn, ByVal Structure As Dom_SchoolStructure) As String()
    Dim Table() As String
    ReDim Table(1 To Structure.PeriodCount, 1 To Structure.MainStreamGradeCount)
    Dim G As Long
    For G = 1 To Structure.MainStreamGradeCount
        Dim Grade As Dom_Grade
        Set Grade = CreateGrade(G)
        Dim P As Long
        For P = 1 To Structure.PeriodCount
            Dim Period As Dom_Period
            Set Period = CreatePeriod(P)
            Dim Filter As Dom_ClassHourFilter
            Set Filter = CreateFilter(ClassHourType, Period, Grade)
            Table(P, G) = ClassHour.GetMarkCount(Filter)
        Next
    Next
    ToTimeTable = Table
End Function

Private Function CreateFilter(ByVal ClassHourTypeEnum As Dom_ClassHourTypeEnum, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_ClassHourFilter
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = CreateClassHourType(ClassHourTypeEnum)
    With New Dom_ClassHourFilter
        Set CreateFilter = .Create(ClassHourType, Period, Grade)
    End With
End Function

Private Function CreateClassHourType(ByVal Value As Long) As Dom_ClassHourType
    With New Dom_ClassHourType
        Set CreateClassHourType = .Create(Value)
    End With
End Function

Private Function CreatePeriod(ByVal Value As Long) As Dom_Period
    With New Dom_Period
        Set CreatePeriod = .Create(Value)
    End With
End Function

Private Function CreateGrade(ByVal Value As Long) As Dom_Grade
    With New Dom_Grade
        Set CreateGrade = .Create(Value)
    End With
End Function
