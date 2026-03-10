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
    LimitValue      As Dom_LimitValue
    Subject         As Dom_Subject
    TargetDate      As Date
    Execution       As ExecutionMember
    Plan            As PlanMember
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal LimitValue As Dom_LimitValue, ByVal Subject As Dom_Subject, ByVal TargetDate As Date)
    Set This.LimitValue = LimitValue
    Set This.Subject = Subject
    This.TargetDate = TargetDate
    Dim SubjectCount As Long
    SubjectCount = Subject.Count
    Dim PeriodCount As Long
    PeriodCount = LimitValue.PeriodCount
    Dim GradeCount As Long
    GradeCount = LimitValue.MainStreamGradeCount
    ReDim This.Execution.ClassHourSubject(1 To SubjectCount, 1 To GradeCount)
    ReDim This.Execution.TimeTablePeriod(1 To PeriodCount, 1 To GradeCount)
    ReDim This.Plan.ClassHourSubject(1 To SubjectCount, 1 To GradeCount)
    ReDim This.Plan.TimeTablePeriod(1 To PeriodCount, 1 To GradeCount)
End Sub

Public Property Get LimitValue() As Dom_LimitValue
    Set LimitValue = This.LimitValue
End Property

Public Property Get TargetDate() As Date
    TargetDate = This.TargetDate
End Property

Public Sub AddExecutionClassHourSubject(ByVal Index As Long, ByVal Grade As Long, ByVal Value As Long)
    This.Execution.ClassHourSubject(Index, Grade) = This.Execution.ClassHourSubject(Index, Grade) + Value
End Sub

Public Sub AddPlanClassHourSubject(ByVal Index As Long, ByVal Grade As Long, ByVal Value As Long)
    This.Plan.ClassHourSubject(Index, Grade) = This.Plan.ClassHourSubject(Index, Grade) + Value
End Sub

Public Sub AddExecutionTimeTablePeriod(ByVal Period As Long, ByVal Grade As Long, ByVal Value As Long)
    This.Execution.TimeTablePeriod(Period, Grade) = This.Execution.TimeTablePeriod(Period, Grade) + Value
End Sub

Public Sub AddPlanTimeTablePeriod(ByVal Period As Long, ByVal Grade As Long, ByVal Value As Long)
    This.Plan.TimeTablePeriod(Period, Grade) = This.Plan.TimeTablePeriod(Period, Grade) + Value
End Sub

Public Function GetExecutionClassHourOf(ByVal Index As Long, ByVal Grade As Long) As Long
    GetExecutionClassHourOf = This.Execution.ClassHourSubject(Index, Grade)
End Function

Public Function GetPlanClassHourOf(ByVal Index As Long, ByVal Grade As Long) As Long
    GetPlanClassHourOf = This.Plan.ClassHourSubject(Index, Grade)
End Function

Public Function GetExecutionTimeTableOf(ByVal Period As Long, ByVal Grade As Long) As Long
    GetExecutionTimeTableOf = This.Execution.TimeTablePeriod(Period, Grade)
End Function

Public Function GetPlanTimeTableOf(ByVal Period As Long, ByVal Grade As Long) As Long
    GetPlanTimeTableOf = This.Plan.TimeTablePeriod(Period, Grade)
End Function
