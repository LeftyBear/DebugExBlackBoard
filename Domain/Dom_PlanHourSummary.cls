VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_PlanHourSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    LowerGrade As Long
    UpperGrade As Long
    LowerSubject As Long
    UpperSubject As Long
    LowerPeriod As Long
    UpperPeriod As Long
    ClassHourTotal() As Long
    TimeTableTotal() As Long
    IsBuilt As Boolean
End Type

Private This As Member

Friend Sub Build(ByVal Entities As VBA.Collection)
    If This.IsBuilt Then Err.Raise 5
    Dim First As Dom_ClassHour
    Set First = Entities.Item(1)
    This.LowerGrade = First.LowerGrade
    This.UpperGrade = First.UpperGrade
    This.LowerSubject = First.LowerSubject
    This.UpperSubject = First.UpperSubject
    This.LowerPeriod = First.LowerPeriod
    This.UpperPeriod = First.UpperPeriod
    ReDim This.ClassHourTotal(This.LowerGrade To This.UpperGrade, This.LowerSubject To This.UpperSubject)
    ReDim This.TimeTableTotal(This.LowerGrade To This.UpperGrade, This.LowerPeriod To This.UpperPeriod)
    Dim Entity As Dom_ClassHour
    For Each Entity In Entities
        Dim Grade As Long
        For Grade = This.LowerGrade To This.UpperGrade
            This.ClassHourTotal(Grade) = Entity.SumPlanClassHourSubject(Grade)
            This.TimeTableTotal(Grade) = Entity.SumPlanTimeTablePeriod(Grade)
        Next
    Next
    This.IsBuilt = True
End Sub

Public Function ClassHourTotal(ByVal Grade As Long, ByVal Subject As Long) As Long
    ClassHourTotal = This.ClassHourTotal(Grade, Subject)
End Function

Public Function TimeTableTotal(ByVal Grade As Long, ByVal Period As Long) As Long
    TimeTableTotal = This.TimeTableTotal(Grade, Period)
End Function

