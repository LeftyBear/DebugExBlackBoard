VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.VO"
Option Explicit

Private Type ExeClassHourMember
    HourTotal()     As Long
End Type
Private Type PlanClassHourMember
    HourTotal()     As Long
End Type
Private Type ExeTimeTableMember
    HourTotal()     As Long
End Type
Private Type PlanTimeTableMember
    HourTotal()     As Long
End Type
Private Type ExecutionMember
    ClassHour       As ExeClassHourMember
    TimeTable       As ExeTimeTableMember
End Type
Private Type PlanMember
    ClassHour       As PlanClassHourMember
    TimeTable       As PlanTimeTableMember
End Type
Private Type Member
    Entities        As VBA.Collection
    LowerGrade      As Long
    UpperGrade      As Long
    LowerSubject    As Long
    UpperSubject    As Long
    LowerPeriod     As Long
    UpperPeriod     As Long
    Exe             As ExecutionMember
    Plan            As PlanMember
End Type
Private This As Member

Friend Sub Inject(ByVal Entities As VBA.Collection)
    If Entities Is Nothing Then
        Err.Raise Dom_ErrNum.NothingEntities, , "Entities is Nothing"
    End If
    If Entities.Count = 0 Then
        Err.Raise Dom_ErrNum.EmptyEntities, , "Entities is empty"
    End If
    Set This.Entities = Entities
    Dim First As Dom_ClassHour
    Set First = Entities.Item(1)
    This.LowerGrade = First.LowerGrade
    This.UpperGrade = First.UpperGrade
    This.LowerSubject = First.LowerSubject
    This.UpperSubject = First.UpperSubject
    This.LowerPeriod = First.LowerPeriod
    This.UpperPeriod = First.UpperPeriod
    ReDim This.Exe.ClassHour.HourTotal(First.LowerGrade To First.UpperGrade)
    ReDim This.Plan.ClassHour.HourTotal(First.LowerGrade To First.UpperGrade)
    ReDim This.Exe.TimeTable.HourTotal(First.LowerGrade To First.UpperGrade)
    ReDim This.Plan.TimeTable.HourTotal(First.LowerGrade To First.UpperGrade)
    AggregateEntities
End Sub

Public Function LowerGrade() As Long
    LowerGrade = This.LowerGrade
End Function

Public Function UpperGrade() As Long
    UpperGrade = This.UpperGrade
End Function

Public Function LowerSubject() As Long
    LowerSubject = This.LowerSubject
End Function

Public Function UpperSubject() As Long
    UpperSubject = This.UpperSubject
End Function

Public Function LowerPeriod() As Long
    LowerPeriod = This.LowerPeriod
End Function

Public Function UpperPeriod() As Long
    UpperPeriod = This.UpperPeriod
End Function

Public Function ExeClassHourTotal(ByVal Grade As Long) As Long
    ExeClassHourTotal = This.Exe.ClassHour.HourTotal(Grade)
End Function

Public Function PlanClassHourTotal(ByVal Grade As Long) As Long
    PlanClassHourTotal = This.Plan.ClassHour.HourTotal(Grade)
End Function

Public Function ExeTimeTableTotal(ByVal Grade As Long) As Long
    ExeTimeTableTotal = This.Exe.TimeTable.HourTotal(Grade)
End Function

Public Function PlanTimeTableTotal(ByVal Grade As Long) As Long
    PlanTimeTableTotal = This.Plan.TimeTable.HourTotal(Grade)
End Function

Public Function ExeClassHourTotalOfSubject(ByVal SubjectName As String, ByVal Grade As Long) As Long
    ExeClassHourTotalOfSubject = TotalFrom(This.Exe.ClassHour.HourTotal, SubjectName, Grade)
End Function

Public Function PlanClassHourTotalOfSubject(ByVal SubjectName As String, ByVal Grade As Long) As Long
    PlanClassHourTotalOfSubject = TotalFrom(This.Plan.ClassHour.HourTotal, SubjectName, Grade)
End Function

Public Function ExeTimeTableTotalOfSubject(ByVal SubjectName As String, ByVal Grade As Long) As Long
    ExeTimeTableTotalOfSubject = TotalFrom(This.Exe.TimeTable.HourTotal, SubjectName, Grade)
End Function

Public Function PlanTimeTableTotalOfSubject(ByVal SubjectName As String, ByVal Grade As Long) As Long
    PlanTimeTableTotalOfSubject = TotalFrom(This.Plan.TimeTable.HourTotal, SubjectName, Grade)
End Function

Private Function TotalFrom(ByRef Values() As Long, ByVal SubjectName As String, ByVal Grade As Long) As Long
    Dim R As Long
    For R = LBound(Values, D1) To UBound(Values, D1)
        Dim Result As Long
        If VBA.Trim$(Values(R, Grade)) = SubjectName Then Result = Result + 1
    Next
    TotalFrom = Result
End Function

Private Sub AggregateEntities()
    Dim Entity As Dom_ClassHour
    For Each Entity In This.Entities
        AggregateClassHour Entity
        AggregateTimeTable Entity
    Next
End Sub

Private Sub AggregateClassHour(ByVal Entity As Dom_ClassHour)
    With This
        Dim Grade As Long
        For Grade = .LowerGrade To .UpperGrade
            .Exe.ClassHour.HourTotal(Grade) = Entity.SumExeClassHourSubject(Grade)
            .Plan.ClassHour.HourTotal(Grade) = Entity.SumPlanClassHourSubject(Grade)
        Next
    End With
End Sub

Private Sub AggregateTimeTable(ByVal Entity As Dom_ClassHour)
    With This
        Dim Grade As Long
        For Grade = .LowerGrade To .UpperGrade
            .Exe.TimeTable.HourTotal(Grade) = Entity.SumExeTimeTablePeriod(Grade)
            .Plan.TimeTable.HourTotal(Grade) = Entity.SumPlanTimeTablePeriod(Grade)
        Next
    End With
End Sub
