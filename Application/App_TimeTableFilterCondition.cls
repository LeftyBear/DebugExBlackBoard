VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableFilterCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.ValueObject")
Option Explicit

Private Type Member
    ClassHourType   As Dom_ClassHourType
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    Set This.ClassHourType = ClassHourType
    Set This.Period = Period
    Set This.Grade = Grade
End Sub

Public Property Get HasType() As Boolean
    If Not This.ClassHourType Is Nothing Then HasType = True
End Property

Public Property Get HasPeriod() As Boolean
    If Not This.Period Is Nothing Then HasPeriod = True
End Property

Public Property Get HasGrade() As Boolean
    If Not This.Grade Is Nothing Then HasGrade = True
End Property

Public Function PeriodMatchs(ByVal Item As Dom_TimeTable) As Boolean
    If Not MatchClassHourType(Item) Then Exit Function
    If Not MatchTimeTablePeriod(Item) Then Exit Function
    PeriodMatchs = True
End Function

Public Function GradeMatchs(ByVal Item As Dom_TimeTable) As Boolean
    If Not MatchClassHourType(Item) Then Exit Function
    If Not MatchTimeTableGrade(Item) Then Exit Function
    GradeMatchs = True
End Function

Private Function MatchClassHourType(ByVal Item As Dom_TimeTable) As Boolean
    If Not Me.HasType Then Exit Function
    If Not This.ClassHourType.Equals(Item.ID.ClassHourType) Then Exit Function
    MatchClassHourType = True
End Function

Private Function MatchTimeTablePeriod(ByVal Item As Dom_TimeTable) As Boolean
    If Not Me.HasPeriod Then Exit Function
    If Not This.Period.Equals(Item.ID.Period) Then Exit Function
    MatchTimeTablePeriod = True
End Function

Private Function MatchTimeTableGrade(ByVal Item As Dom_TimeTable) As Boolean
    If Not Me.HasGrade Then Exit Function
    If Not This.Grade.Equals(Item.ID.Grade) Then Exit Function
    MatchTimeTableGrade = True
End Function
