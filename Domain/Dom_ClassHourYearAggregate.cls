VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourYearAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type PlanGradeMember
    CumulativeTotal()   As Long
    DailyTotal()        As Long
    WeeklyTotal()       As Long
    MonthlyTotal()      As Long
End Type

Private Type PlanSubjectMember
    CumulativeTotal()   As Long
    DailyTotal()        As Long
    WeeklyTotal()       As Long
    MonthlyTotal()      As Long
End Type

Private Type ExecutionGradeMember
    CumulativeTotal()   As Long
    DailyTotal()        As Long
    WeeklyTotal()       As Long
    MonthlyTotal()      As Long
End Type

Private Type ExecutionSubjectMember
    CumulativeTotal()   As Long
    DailyTotal()        As Long
    WeeklyTotal()       As Long
    MonthlyTotal()      As Long
End Type

Private Type Member
    Map         As Scripting.Dictionary
    PlanGrade   As PlanGradeMember
    PlanSubject As PlanSubjectMember
    ExeGrade    As ExecutionGradeMember
    ExeSubject  As ExecutionSubjectMember
End Type

Private This As Member

Public Sub Initialize(ByVal SubjectCount As Long, ByVal GradeCount As Long)
    Set This.Map = New Scripting.Dictionary
    ReDim This.PlanGrade.CumulativeTotal(1 To 366, 1 To GradeCount)
    ReDim This.PlanGrade.DailyTotal(1 To 366, 1 To GradeCount)
    ReDim This.PlanGrade.WeeklyTotal(1 To 53, 1 To GradeCount)
    ReDim This.PlanGrade.MonthlyTotal(1 To 12, 1 To GradeCount)
    ReDim This.PlanSubject.CumulativeTotal(1 To 366, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.PlanSubject.DailyTotal(1 To 366, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.PlanSubject.WeeklyTotal(1 To 53, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.PlanSubject.MonthlyTotal(1 To 12, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.ExeGrade.CumulativeTotal(1 To 366, 1 To GradeCount)
    ReDim This.ExeGrade.DailyTotal(1 To 366, 1 To GradeCount)
    ReDim This.ExeGrade.WeeklyTotal(1 To 53, 1 To GradeCount)
    ReDim This.ExeGrade.MonthlyTotal(1 To 12, 1 To GradeCount)
    ReDim This.ExeSubject.CumulativeTotal(1 To 366, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.ExeSubject.DailyTotal(1 To 366, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.ExeSubject.WeeklyTotal(1 To 53, 1 To SubjectCount, 1 To GradeCount)
    ReDim This.ExeSubject.MonthlyTotal(1 To 12, 1 To SubjectCount, 1 To GradeCount)
End Sub

Public Sub Add(ByVal TargetDate As Date, ByVal Aggregate As Dom_ClassHourAggregate)
    This.Map.Add TargetDate, Aggregate
    SumSubject TargetDate, Aggregate
    SumSubjectOfGrade TargetDate, Aggregate
End Sub

Public Sub UpdateHours(ByVal TargetDate As Date, ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade, ByVal NewValue As Long)
    Dim Aggregate As Dom_ClassHourAggregate
    Set Aggregate = This.Map.Item(TargetDate)
    Dim Diff As Long
'    Diff = Aggregate.UpdateHours(ClassHourType, Subject, Grade, NewValue)
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    Dim WeekCount As Long
    WeekCount = Dom_Util_Date.GetSchoolWeek(TargetDate)
    Dim MonthCount As Long
    MonthCount = VBA.Month(TargetDate)
    Select Case ClassHourType
    Case Dom_ClassHourTypeEnum.Plan
        This.PlanGrade.CumulativeTotal(DayCount, Grade.Value) = This.PlanGrade.CumulativeTotal(DayCount, Grade.Value) + Diff
        This.PlanGrade.DailyTotal(DayCount, Grade.Value) = This.PlanGrade.DailyTotal(DayCount, Grade.Value) + Diff
        This.PlanGrade.WeeklyTotal(WeekCount, Grade.Value) = This.PlanGrade.WeeklyTotal(WeekCount, Grade.Value) + Diff
        This.PlanGrade.MonthlyTotal(MonthCount, Grade.Value) = This.PlanGrade.MonthlyTotal(MonthCount, Grade.Value) + Diff
    Case Dom_ClassHourTypeEnum.Execution
        This.ExeGrade.CumulativeTotal(DayCount, Grade.Value) = This.ExeGrade.CumulativeTotal(DayCount, Grade.Value) + Diff
        This.ExeGrade.DailyTotal(DayCount, Grade.Value) = This.ExeGrade.DailyTotal(DayCount, Grade.Value) + Diff
        This.ExeGrade.WeeklyTotal(WeekCount, Grade.Value) = This.ExeGrade.WeeklyTotal(WeekCount, Grade.Value) + Diff
        This.ExeGrade.MonthlyTotal(MonthCount, Grade.Value) = This.ExeGrade.MonthlyTotal(MonthCount, Grade.Value) + Diff
    End Select
End Sub

Public Function GetAggregate(ByVal TargetDate As Date) As Dom_ClassHourAggregate
    Set GetAggregate = This.Map.Item(TargetDate)
End Function

Public Function GetPlanGradeGrandTotal(ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = UBound(This.PlanGrade.CumulativeTotal)
    GetPlanGradeGrandTotal = This.PlanGrade.CumulativeTotal(DayCount, Grade.Value)
End Function

Public Function GetPlanGradeCumulativeTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetPlanGradeCumulativeTotal = This.PlanGrade.CumulativeTotal(DayCount, Grade.Value)
End Function

Public Function GetPlanGradeDailyTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetPlanGradeDailyTotal = This.PlanGrade.DailyTotal(DayCount, Grade.Value)
End Function

Public Function GetPlanGradeWeeklyTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim WeekCount As Long
    WeekCount = Dom_Util_Date.GetSchoolWeek(TargetDate)
    GetPlanGradeWeeklyTotal = This.PlanGrade.WeeklyTotal(WeekCount, Grade.Value)
End Function

Public Function GetPlanGradeMonthlyTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim MonthCount As Long
    MonthCount = VBA.Month(TargetDate)
    GetPlanGradeMonthlyTotal = This.PlanGrade.MonthlyTotal(MonthCount, Grade.Value)
End Function

Public Function GetPlanSubjectCumulativeTotal(ByVal TargetDate As Date, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetPlanSubjectCumulativeTotal = This.PlanSubject.CumulativeTotal(DayCount, Subject.Index, Grade.Value)
End Function

Public Function GetPlanSubjectDailyTotal(ByVal TargetDate As Date, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetPlanSubjectDailyTotal = This.PlanSubject.DailyTotal(DayCount, Subject.Index, Grade.Value)
End Function

Public Function GetPlanSubjectWeeklyTotal(ByVal TargetDate As Date, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetPlanSubjectWeeklyTotal = This.PlanSubject.WeeklyTotal(DayCount, Subject.Index, Grade.Value)
End Function

Public Function GetPlanSubjectMonthlyTotal(ByVal TargetDate As Date, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetPlanSubjectMonthlyTotal = This.PlanSubject.MonthlyTotal(DayCount, Subject.Index, Grade.Value)
End Function

Public Function GetExeGradeCumulativeTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetExeGradeCumulativeTotal = This.ExeGrade.CumulativeTotal(DayCount, Grade.Value)
End Function

Public Function GetExeGradeDailyTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetExeGradeDailyTotal = This.ExeGrade.DailyTotal(DayCount, Grade.Value)
End Function

Public Function GetExeGradeWeeklyTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim WeekCount As Long
    WeekCount = Dom_Util_Date.GetSchoolWeek(TargetDate)
    GetExeGradeWeeklyTotal = This.ExeGrade.WeeklyTotal(WeekCount, Grade.Value)
End Function

Public Function GetExeGradeMonthlyTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim MonthCount As Long
    MonthCount = VBA.Month(TargetDate)
    GetExeGradeMonthlyTotal = This.ExeGrade.MonthlyTotal(MonthCount, Grade.Value)
End Function

Private Sub SumSubject(ByVal TargetDate As Date, ByVal Aggregate As Dom_ClassHourAggregate)
'    Dim DayCount As Long
'    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
'    Dim WeekCount As Long
'    WeekCount = Dom_Util_Date.GetSchoolWeek(TargetDate)
'    Dim MonthCount As Long
'    MonthCount = VBA.Month(TargetDate)
'    Dim i As Long
'    For i = 1 To Aggregate.SubjectCount
'        Dim G As Long
'        For G = 1 To Aggregate.GradeCount
'            Dim Grade As Dom_Grade
'            Set Grade = CreateGrade(G)
'            Dim PlanGradeSum As Long
'            PlanGradeSum = PlanGradeSum + Aggregate.GetHours(Dom_ClassHourType.Plan, CreateSubject(i), Grade)
'            Dim ExeGradeSum As Long
'            ExeGradeSum = ExeGradeSum + Aggregate.GetHours(Dom_ClassHourType.Execution, CreateSubject(i), Grade)
'            This.PlanSubject.CumulativeTotal(DayCount, i, G) = This.PlanSubject.CumulativeTotal(DayCount, i, G) + PlanGradeSum
'            This.PlanSubject.DailyTotal(DayCount, i, G) = PlanGradeSum
'            This.PlanSubject.WeeklyTotal(WeekCount, i, G) = This.PlanSubject.WeeklyTotal(WeekCount, i, G) + PlanGradeSum
'            This.PlanSubject.MonthlyTotal(MonthCount, i, G) = This.PlanSubject.MonthlyTotal(MonthCount, i, G) + PlanGradeSum
'            This.ExeSubject.CumulativeTotal(DayCount, i, G) = This.ExeSubject.CumulativeTotal(DayCount, i, G) + ExeGradeSum
'            This.ExeSubject.DailyTotal(DayCount, i, G) = ExeGradeSum
'            This.ExeSubject.WeeklyTotal(WeekCount, i, G) = This.ExeSubject.WeeklyTotal(WeekCount, i, G) + ExeGradeSum
'            This.ExeSubject.MonthlyTotal(MonthCount, i, G) = This.ExeSubject.MonthlyTotal(MonthCount, i, G) + ExeGradeSum
'        Next
'    Next
End Sub

Private Sub SumSubjectOfGrade(ByVal TargetDate As Date, ByVal Aggregate As Dom_ClassHourAggregate)
'    Dim DayCount As Long
'    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
'    Dim WeekCount As Long
'    WeekCount = Dom_Util_Date.GetSchoolWeek(TargetDate)
'    Dim MonthCount As Long
'    MonthCount = VBA.Month(TargetDate)
'    Dim G As Long
'    For G = 1 To Aggregate.GradeCount
'        Dim Grade As Dom_Grade
'        Set Grade = CreateGrade(G)
'        Dim i As Long
'        For i = 1 To Aggregate.SubjectCount
'            Dim PlanSubjectSum As Long
'            PlanSubjectSum = PlanSubjectSum + Aggregate.GetHours(Dom_ClassHourType.Plan, CreateSubject(i), Grade)
'            Dim ExeSubjectSum As Long
'            ExeSubjectSum = ExeSubjectSum + Aggregate.GetHours(Dom_ClassHourType.Execution, CreateSubject(i), Grade)
'        Next
'        This.PlanGrade.CumulativeTotal(DayCount, G) = This.PlanGrade.CumulativeTotal(DayCount, G) + PlanSubjectSum
'        This.PlanGrade.DailyTotal(DayCount, G) = PlanSubjectSum
'        This.PlanGrade.WeeklyTotal(WeekCount, G) = This.PlanGrade.WeeklyTotal(WeekCount, G) + PlanSubjectSum
'        This.PlanGrade.MonthlyTotal(MonthCount, G) = This.PlanGrade.MonthlyTotal(MonthCount, G) + PlanSubjectSum
'        This.ExeGrade.CumulativeTotal(DayCount, G) = This.ExeGrade.CumulativeTotal(DayCount, G) + ExeSubjectSum
'        This.ExeGrade.DailyTotal(DayCount, G) = ExeSubjectSum
'        This.ExeGrade.WeeklyTotal(WeekCount, G) = This.ExeGrade.WeeklyTotal(WeekCount, G) + ExeSubjectSum
'        This.ExeGrade.MonthlyTotal(MonthCount, G) = This.ExeGrade.MonthlyTotal(MonthCount, G) + ExeSubjectSum
'    Next
End Sub

Private Function CreateSubject(ByVal Index As Long) As Dom_Subject
    With New Dom_Subject
        Set CreateSubject = .Create(Index, vbNullString, vbNullString, vbNullString)
    End With
End Function

Private Function CreateGrade(ByVal Value As Long) As Dom_Grade
    With New Dom_Grade
        Set CreateGrade = .Create(Value)
    End With
End Function
