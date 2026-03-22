VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    HasType         As Boolean
    HasSubject      As Boolean
    HasPeriod       As Boolean
    HasGrade        As Boolean
    ClassHourType   As Dom_ClassHourType
    Subject         As Dom_Subject
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    If Not ClassHourType Is Nothing Then Set This.ClassHourType = ClassHourType: This.HasType = True
    If Not Subject Is Nothing Then Set This.Subject = Subject: This.HasSubject = True
    If Not Period Is Nothing Then Set This.Period = Period: This.HasPeriod = True
    If Not Grade Is Nothing Then Set This.Grade = Grade: This.HasGrade = True
End Sub

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_ClassHourFilter
    Dim Result As Dom_ClassHourFilter
    Set Result = New Dom_ClassHourFilter
    Result.InitializeInternal ClassHourType, Subject, Period, Grade
    Set Create = Result
End Function

Public Function IsMatch(ByVal Item As Dom_ClassHour) As Boolean
    If Not MatchType(Item) Then Exit Function
    If Not MatchSubject(Item) Then Exit Function
    If Not MatchPeriod(Item) Then Exit Function
    If Not MatchGrade(Item) Then Exit Function
    IsMatch = True
End Function

Public Property Get HasType() As Boolean
    HasType = This.HasType
End Property

Public Property Get HasSubject() As Boolean
    HasSubject = This.HasSubject
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

Public Property Get Subject() As Dom_Subject
    Set Subject = This.Subject
End Property

Public Property Get Period() As Dom_Period
    Set Period = This.Period
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Private Function MatchType(ByVal Item As Dom_ClassHour) As Boolean
    If Not This.HasType Then Exit Function
    If Not This.ClassHourType.Equals(Item.GetID.ClassHourType) Then Exit Function
    MatchType = True
End Function

Private Function MatchSubject(ByVal Item As Dom_ClassHour) As Boolean
    If Not This.HasSubject Then Exit Function
    If Not This.Subject.Equals(Item.GetID.Subject) Then Exit Function
    MatchSubject = True
End Function

Private Function MatchPeriod(ByVal Item As Dom_ClassHour) As Boolean
    If Not This.HasPeriod Then Exit Function
    If Not This.Period.Equals(Item.GetID.Period) Then Exit Function
    MatchPeriod = True
End Function

Private Function MatchGrade(ByVal Item As Dom_ClassHour) As Boolean
    If Not This.HasGrade Then Exit Function
    If Not This.Grade.Equals(Item.GetID.Grade) Then Exit Function
    MatchGrade = True
End Function
