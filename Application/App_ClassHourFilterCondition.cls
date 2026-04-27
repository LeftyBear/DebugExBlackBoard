VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourFilterCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.ValueObject")
Option Explicit

Private Type Member
    ClassHourType   As Dom_ClassHourType
    Subject         As Dom_Subject
    Grade           As Dom_Grade
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade)
    Set This.ClassHourType = ClassHourType
    Set This.Subject = Subject
    Set This.Grade = Grade
End Sub

Public Property Get HasType() As Boolean
    If Not This.ClassHourType Is Nothing Then HasType = True
End Property

Public Property Get HasSubject() As Boolean
    If Not This.Subject Is Nothing Then HasSubject = True
End Property

Public Property Get HasGrade() As Boolean
    If Not This.Grade Is Nothing Then HasGrade = True
End Property

Public Function SubjectMatchs(ByVal Item As Dom_ClassHour) As Boolean
    If Not MatchClassHourType(Item) Then Exit Function
    If Not MatchClassHourSubject(Item) Then Exit Function
    SubjectMatchs = True
End Function

Public Function GradeMatchs(ByVal Item As Dom_ClassHour) As Boolean
    If Not MatchClassHourType(Item) Then Exit Function
    If Not MatchClassHourGrade(Item) Then Exit Function
    GradeMatchs = True
End Function

Private Function MatchClassHourType(ByVal Item As Dom_ClassHour) As Boolean
    If Not Me.HasType Then Exit Function
    If Not This.ClassHourType.Equals(Item.ID.ClassHourType) Then Exit Function
    MatchClassHourType = True
End Function

Private Function MatchClassHourSubject(ByVal Item As Dom_ClassHour) As Boolean
    If Not Me.HasSubject Then Exit Function
    If Not This.Subject.Equals(Item.ID.Subject) Then Exit Function
    MatchClassHourSubject = True
End Function

Private Function MatchClassHourGrade(ByVal Item As Dom_ClassHour) As Boolean
    If Not Me.HasGrade Then Exit Function
    If Not This.Grade.Equals(Item.ID.Grade) Then Exit Function
    MatchClassHourGrade = True
End Function
