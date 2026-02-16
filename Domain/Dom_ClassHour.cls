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

Private Type ExeClassHourMember
    Subject()   As Long
End Type

Private Type PlanClassHourMember
    Subject()   As Long
End Type

Private Type ExeTimeTableMember
    Period()    As String
End Type

Private Type PlanTimeTableMember
    Period()    As String
End Type

Private Type ExecutionMember
    ClassHour   As ExeClassHourMember
    TimeTable   As ExeTimeTableMember
End Type

Private Type PlanMember
    ClassHour   As PlanClassHourMember
    TimeTable   As PlanTimeTableMember
End Type

Private Type Member
    TargetDate  As Date
    Exe         As ExecutionMember
    Plan        As PlanMember
End Type

Private This As Member

Friend Sub Initialize(ByVal TargetDate As Date, ByRef ExeClassHourSubject() As Long, ByRef ExeTimeTablePeriod() As String, ByRef PlanClassHourSubject() As Long, ByRef PlanTimeTablePeriod() As String)
    This.TargetDate = TargetDate
    ReDim This.Exe.ClassHour.Subject(LBound(ExeClassHourSubject, D1) To UBound(ExeClassHourSubject, D1), LBound(ExeClassHourSubject, D2) To UBound(ExeClassHourSubject, D2))
    ReDim This.Exe.TimeTable.Period(LBound(ExeTimeTablePeriod, D1) To UBound(ExeTimeTablePeriod, D1), LBound(ExeTimeTablePeriod, D2) To UBound(ExeTimeTablePeriod, D2))
    ReDim This.Plan.ClassHour.Subject(LBound(PlanClassHourSubject, D1) To UBound(PlanClassHourSubject, D1), LBound(PlanClassHourSubject, D2) To UBound(PlanClassHourSubject, D2))
    ReDim This.Plan.TimeTable.Period(LBound(PlanTimeTablePeriod, D1) To UBound(PlanTimeTablePeriod, D1), LBound(PlanTimeTablePeriod, D2) To UBound(PlanTimeTablePeriod, D2))
    This.Exe.ClassHour.Subject = CopyLongArray(ExeClassHourSubject)
    This.Exe.TimeTable.Period = CopyStringArray(ExeTimeTablePeriod)
    This.Plan.ClassHour.Subject = CopyLongArray(PlanClassHourSubject)
    This.Plan.TimeTable.Period = CopyStringArray(PlanTimeTablePeriod)
End Sub

Public Property Get TargetDate() As Date
    TargetDate = This.TargetDate
End Property

Public Property Get UpperGrade() As Long
    UpperGrade = UBound(This.Exe.ClassHour.Subject, D2)
End Property

Public Property Get LowerGrade() As Long
    LowerGrade = LBound(This.Exe.ClassHour.Subject, D2)
End Property

Public Property Get UpperSubject() As Long
    UpperSubject = UBound(This.Exe.ClassHour.Subject, D1)
End Property

Public Property Get LowerSubject() As Long
    LowerSubject = LBound(This.Exe.ClassHour.Subject, D1)
End Property

Public Property Get UpperPeriod() As Long
    UpperPeriod = UBound(This.Exe.TimeTable.Period, D1)
End Property

Public Property Get LowerPeriod() As Long
    LowerPeriod = LBound(This.Exe.TimeTable.Period, D1)
End Property

Public Function ExeClassHourSubject(ByVal SubjectCount As Long, ByVal Grade As Long) As Long
    ValidateSubjectCount SubjectCount
    ValidateGrade Grade
    ExeClassHourSubject = This.Exe.ClassHour.Subject(SubjectCount, Grade)
End Function

Public Function ExeTimeTablePeriod(ByVal PeriodCount As Long, ByVal Grade As Long) As String
    ValidatePeriodCount PeriodCount
    ValidateGrade Grade
    ExeTimeTablePeriod = This.Exe.TimeTable.Period(PeriodCount, Grade)
End Function

Public Function SumExeClassHourSubject(ByVal Grade As Long) As Long
    ValidateSubjectCount Grade
    SumExeClassHourSubject = SumClassHour(This.Exe.ClassHour.Subject, Grade)
End Function

Public Function SumExeTimeTablePeriod(ByVal Grade As Long) As Long
    ValidateSubjectCount Grade
    SumExeTimeTablePeriod = ConutSubject(This.Exe.TimeTable.Period, Grade)
End Function

Public Function PlanClassHourSubject(ByVal SubjectCount As Long, ByVal Grade As Long) As Long
    ValidateGrade Grade
    ValidateSubjectCount SubjectCount
    PlanClassHourSubject = This.Plan.ClassHour.Subject(SubjectCount, Grade)
End Function

Public Function PlanTimeTablePeriod(ByVal PeriodCount As Long, ByVal Grade As Long) As String
    ValidateGrade Grade
    ValidatePeriodCount PeriodCount
    PlanTimeTablePeriod = This.Plan.TimeTable.Period(PeriodCount, Grade)
End Function

Public Function SumPlanClassHourSubject(ByVal Grade As Long) As Long
    ValidateGrade Grade
    SumPlanClassHourSubject = SumClassHour(This.Plan.ClassHour.Subject, Grade)
End Function

Public Function SumPlanTimeTablePeriod(ByVal Grade As Long) As Long
    ValidateGrade Grade
    SumPlanTimeTablePeriod = ConutSubject(This.Plan.TimeTable.Period, Grade)
End Function

Private Function SumClassHour(ByRef Values() As Long, ByVal Grade As Long) As Long
    Dim R As Long
    For R = LBound(Values, D1) To UBound(Values, D1)
        Dim Result As Long
        Result = Result + Values(R, Grade)
    Next
    SumClassHour = Result
End Function

Private Function ConutSubject(ByRef Values() As String, ByVal Grade As Long) As Long
    Dim R As Long
    For R = LBound(Values, D1) To UBound(Values, D1)
        Dim Result As Long
        If VBA.Trim$(Values(R, Grade)) <> vbNullString Then Result = Result + 1
    Next
    ConutSubject = Result
End Function

Private Sub ValidateGrade(ByVal Grade As Long)
    If Grade < LBound(This.Exe.ClassHour.Subject, D2) Or UBound(This.Exe.ClassHour.Subject, D2) < Grade Then
        Err.Raise Dom_ErrNum.OutOfRange, , "Grade is out of range"
    End If
End Sub

Private Sub ValidateSubjectCount(ByVal SubjectCount As Long)
    If SubjectCount < LBound(This.Exe.ClassHour.Subject, D1) Or UBound(This.Exe.ClassHour.Subject, D1) < SubjectCount Then
        Err.Raise Dom_ErrNum.OutOfRange, , "SubjectCount is out of range"
    End If
End Sub

Private Sub ValidatePeriodCount(ByVal PeriodCount As Long)
    If PeriodCount < LBound(This.Exe.TimeTable.Period, D1) Or UBound(This.Exe.TimeTable.Period, D1) < PeriodCount Then
        Err.Raise Dom_ErrNum.OutOfRange, , "PeriodCount is out of range"
    End If
End Sub

Private Function CopyLongArray(ByRef Values() As Long) As Long()
    Dim Result() As Long
    ReDim Result(LBound(Values, D1) To UBound(Values, D1), LBound(Values, D2) To UBound(Values, D2))
    Dim R As Long
    For R = LBound(Values, D1) To UBound(Values, D1)
        Dim C As Long
        For C = LBound(Values, D2) To UBound(Values, D2)
            Result(R, C) = Values(R, C)
        Next
    Next
    CopyLongArray = Result
End Function

Private Function CopyStringArray(ByRef Values() As String) As String()
    Dim Result() As String
    ReDim Result(LBound(Values, D1) To UBound(Values, D1), LBound(Values, D2) To UBound(Values, D2))
    Dim R As Long
    For R = LBound(Values, D1) To UBound(Values, D1)
        Dim C As Long
        For C = LBound(Values, D2) To UBound(Values, D2)
            Result(R, C) = Values(R, C)
        Next
    Next
    CopyStringArray = Result
End Function

