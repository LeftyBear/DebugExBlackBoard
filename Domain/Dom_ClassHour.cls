VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type ExecutionMember
    ClassHourSubject()  As Long
    TimeTablePeriod()   As String
End Type

Private Type PlanMember
    ClassHourSubject()  As Long
    TimeTablePeriod()   As String
End Type

Private Type Member
    Structure   As Dom_SchoolStructure
    TargetDate  As Date
    Execution   As ExecutionMember
    Plan        As PlanMember
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal Structure As Dom_SchoolStructure, ByVal TargetDate As Date)
    Set This.Structure = Structure
    This.TargetDate = TargetDate
    Dim SubjectCount As Long
    SubjectCount = Structure.SubjectCount
    Dim PeriodCount As Long
    PeriodCount = Structure.PeriodCount
    Dim GradeCount As Long
    GradeCount = Structure.MainStreamGradeCount
    ReDim This.Execution.ClassHourSubject(1 To SubjectCount, 1 To GradeCount)
    ReDim This.Execution.TimeTablePeriod(1 To PeriodCount, 1 To GradeCount)
    ReDim This.Plan.ClassHourSubject(1 To SubjectCount, 1 To GradeCount)
    ReDim This.Plan.TimeTablePeriod(1 To PeriodCount, 1 To GradeCount)
End Sub

Public Property Get Structure() As Dom_SchoolStructure
    Set Structure = This.Structure
End Property

Public Property Get TargetDate() As Date
    TargetDate = This.TargetDate
End Property

Public Sub AddExecutionClassHourSubject(ByVal Subject As Long, ByVal Grade As Long, ByVal Value As Long)
    Dim SubjectIndex As Long
    SubjectIndex = This.Structure.ToSubjectIndex(Subject)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    This.Execution.ClassHourSubject(SubjectIndex, GradeIndex) = This.Execution.ClassHourSubject(SubjectIndex, GradeIndex) + Value
End Sub

Public Sub AddPlanClassHourSubject(ByVal Subject As Long, ByVal Grade As Long, ByVal Value As Long)
    Dim SubjectIndex As Long
    SubjectIndex = This.Structure.ToSubjectIndex(Subject)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    This.Plan.ClassHourSubject(SubjectIndex, GradeIndex) = This.Plan.ClassHourSubject(SubjectIndex, GradeIndex) + Value
End Sub

Public Sub AddExecutionTimeTablePeriod(ByVal Period As Long, ByVal Grade As Long, ByVal Value As Long)
    Dim PeriodIndex As Long
    PeriodIndex = This.Structure.ToPeriodIndex(Period)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    This.Execution.TimeTablePeriod(PeriodIndex, GradeIndex) = This.Execution.TimeTablePeriod(PeriodIndex, GradeIndex) + Value
End Sub

Public Sub AddPlanTimeTablePeriod(ByVal Period As Long, ByVal Grade As Long, ByVal Value As Long)
    Dim PeriodIndex As Long
    PeriodIndex = This.Structure.ToPeriodIndex(Period)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    This.Plan.TimeTablePeriod(PeriodIndex, GradeIndex) = This.Plan.TimeTablePeriod(PeriodIndex, GradeIndex) + Value
End Sub

Public Function GetExecutionClassHourOf(ByVal Subject As Long, ByVal Grade As Long) As Long
    Dim SubjectIndex As Long
    SubjectIndex = This.Structure.ToSubjectIndex(Subject)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    GetExecutionClassHourOf = This.Execution.ClassHourSubject(SubjectIndex, GradeIndex)
End Function

Public Function GetPlanClassHourOf(ByVal Subject As Long, ByVal Grade As Long) As Long
    Dim SubjectIndex As Long
    SubjectIndex = This.Structure.ToSubjectIndex(Subject)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    GetPlanClassHourOf = This.Plan.ClassHourSubject(SubjectIndex, GradeIndex)
End Function

Public Function GetExecutionTimeTableOf(ByVal Period As Long, ByVal Grade As Long) As Long
    Dim PeriodIndex As Long
    PeriodIndex = This.Structure.ToPeriodIndex(Period)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    GetExecutionTimeTableOf = This.Execution.TimeTablePeriod(PeriodIndex, GradeIndex)
End Function

Public Function GetPlanTimeTableOf(ByVal Period As Long, ByVal Grade As Long) As Long
    Dim PeriodIndex As Long
    PeriodIndex = This.Structure.ToPeriodIndex(Period)
    Dim GradeIndex As Long
    GradeIndex = This.Structure.ToSpecialGradeIndex(Grade)
    GetPlanTimeTableOf = This.Plan.TimeTablePeriod(PeriodIndex, GradeIndex)
End Function
