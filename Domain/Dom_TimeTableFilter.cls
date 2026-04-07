VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    HasType         As Boolean
    HasPeriod       As Boolean
    HasGrade        As Boolean
    ClassHourType   As Dom_ClassHourType
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    If Not ClassHourType Is Nothing Then Set This.ClassHourType = ClassHourType: This.HasType = True
    If Not Period Is Nothing Then Set This.Period = Period: This.HasPeriod = True
    If Not Grade Is Nothing Then Set This.Grade = Grade: This.HasGrade = True
End Sub

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_ClassHourFilter
    Dim Result As Dom_ClassHourFilter
    Set Result = New Dom_ClassHourFilter
    Result.InitializeInternal ClassHourType, Period, Grade
    Set Create = Result
End Function

Public Function IsMatch(ByVal Item As Dom_TimeTable) As Boolean
    If Not MatchTimeTableType(Item) Then Exit Function
    If Not MatchTimeTablePeriod(Item) Then Exit Function
    If Not MatchTimeTableGrade(Item) Then Exit Function
    IsMatch = True
End Function

Public Property Get HasType() As Boolean
    HasType = This.HasType
End Property

Public Property Get HasPeriod() As Boolean
    HasPeriod = This.HasPeriod
End Property

Public Property Get HasGrade() As Boolean
    HasGrade = This.HasGrade
End Property

Public Property Get ClassHourType() As Dom_ClassHourType
    Set ClassHourType = This.ClassHourType
End Property

Public Property Get Period() As Dom_Period
    Set Period = This.Period
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Private Function MatchTimeTableType(ByVal Item As Dom_TimeTable) As Boolean
    If Not This.HasType Then Exit Function
    If Not This.ClassHourType.Equals(Item.GetID.ClassHourType) Then Exit Function
    MatchTimeTableType = True
End Function

Private Function MatchTimeTablePeriod(ByVal Item As Dom_TimeTable) As Boolean
    If Not This.HasPeriod Then Exit Function
    If Not This.Period.Equals(Item.GetID.Period) Then Exit Function
    MatchTimeTablePeriod = True
End Function

Private Function MatchTimeTableGrade(ByVal Item As Dom_TimeTable) As Boolean
    If Not This.HasGrade Then Exit Function
    If Not This.Grade.Equals(Item.GetID.Grade) Then Exit Function
    MatchTimeTableGrade = True
End Function
