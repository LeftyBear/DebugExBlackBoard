VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    HasType         As Boolean
    HasGrade        As Boolean
    HasClassNo      As Boolean
    HasGender       As Boolean
    StreamType      As Dom_StreamType
    Grade           As Dom_Grade
    ClassNo         As Dom_ClassNo
    Gender          As Dom_Gender
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender)
    If Not StreamType Is Nothing Then Set This.StreamType = StreamType: This.HasType = True
    If Not Grade Is Nothing Then Set This.Grade = Grade: This.HasGrade = True
    If Not ClassNo Is Nothing Then Set This.ClassNo = ClassNo: This.HasClassNo = True
    If Not Gender Is Nothing Then Set This.Gender = Gender: This.HasGender = True
End Sub

Public Function Create(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender) As Dom_EnrollmentFilter
    Dim Result As Dom_EnrollmentFilter
    Set Result = New Dom_EnrollmentFilter
    Result.InitializeInternal StreamType, Grade, ClassNo, Gender
    Set Create = Result
End Function

Public Function IsMatch(ByVal Item As Dom_Enrollment) As Boolean
    If Not MatchType(Item) Then Exit Function
    If Not MatchGrade(Item) Then Exit Function
    If Not MatchClassNo(Item) Then Exit Function
    If Not MatchGender(Item) Then Exit Function
    IsMatch = True
End Function

Public Property Get HasType() As Boolean
    HasType = This.HasType
End Property

Public Property Get HasGrade() As Boolean
    HasGrade = This.HasGrade
End Property

Public Property Get HasClassNo() As Boolean
    HasClassNo = This.HasClassNo
End Property

Public Property Get HasGender() As Boolean
    HasGender = This.HasGender
End Property

Public Property Get StreamType() As Dom_StreamType
    Set StreamType = This.StreamType
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Public Property Get ClassNo() As Dom_ClassNo
    Set ClassNo = This.ClassNo
End Property

Public Property Get Gender() As Dom_Gender
    Set Gender = This.Gender
End Property

Private Function MatchType(ByVal Item As Dom_Enrollment) As Boolean
    If Not This.HasType Then Exit Function
    If Not This.StreamType.Equals(Item.GetID.StreamType) Then Exit Function
    MatchType = True
End Function

Private Function MatchGrade(ByVal Item As Dom_Enrollment) As Boolean
    If Not This.HasGrade Then Exit Function
    If Not This.Grade.Equals(Item.GetID.Grade) Then Exit Function
    MatchGrade = True
End Function

Private Function MatchClassNo(ByVal Item As Dom_Enrollment) As Boolean
    If Not This.HasClassNo Then Exit Function
    If Not This.ClassNo.Equals(Item.GetID.ClassNo) Then Exit Function
    MatchClassNo = True
End Function

Private Function MatchGender(ByVal Item As Dom_Enrollment) As Boolean
    If Not This.HasGender Then Exit Function
    If Not This.Gender.Equals(Item.GetID.Gender) Then Exit Function
    MatchGender = True
End Function
