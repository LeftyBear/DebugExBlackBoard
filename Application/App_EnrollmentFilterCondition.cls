VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentFilterCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.ValueObject")
Option Explicit

Private Type Member
    StreamType      As Dom_StreamType
    Grade           As Dom_Grade
    ClassNo         As Dom_ClassNo
    Gender          As Dom_Gender
End Type

Private This As Member

Public Sub Initialize(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, Gender As Dom_Gender)
    Set This.StreamType = StreamType
    Set This.Grade = Grade
    Set This.ClassNo = ClassNo
    Set This.Gender = Gender
End Sub

Public Property Get HasType() As Boolean
    HasType = This.StreamType
End Property

Public Property Get HasGrade() As Boolean
    HasGrade = This.Grade
End Property

Public Property Get HasClassNo() As Boolean
    HasClassNo = This.ClassNo
End Property

Public Property Get HasGender() As Boolean
    HasGender = This.Gender
End Property

Public Function GenderMatchs(ByVal Item As Dom_ClassHour) As Boolean
    If Not MatchType(Item) Then Exit Function
    If Not MatchGrade(Item) Then Exit Function
    If Not MatchGender(Item) Then Exit Function
    GenderMatchs = True
End Function

Public Function ClassMatchs(ByVal Item As Dom_ClassHour) As Boolean
    If Not MatchType(Item) Then Exit Function
    If Not MatchGrade(Item) Then Exit Function
    If Not MatchClassNo(Item) Then Exit Function
    ClassMatchs = True
End Function

Private Function MatchType(ByVal Item As Dom_Enrollment) As Boolean
    If Not Me.HasType Then Exit Function
    If Not This.StreamType.Equals(Item.ID.StreamType) Then Exit Function
    MatchType = True
End Function

Private Function MatchGrade(ByVal Item As Dom_Enrollment) As Boolean
    If Not Me.HasGrade Then Exit Function
    If Not This.Grade.Equals(Item.ID.Grade) Then Exit Function
    MatchGrade = True
End Function

Private Function MatchClassNo(ByVal Item As Dom_Enrollment) As Boolean
    If Not Me.HasClassNo Then Exit Function
    If Not This.ClassNo.Equals(Item.ID.ClassNo) Then Exit Function
    MatchClassNo = True
End Function

Private Function MatchGender(ByVal Item As Dom_Enrollment) As Boolean
    If Not Me.HasGender Then Exit Function
    If Not This.Gender.Equals(Item.ID.Gender) Then Exit Function
    MatchGender = True
End Function
