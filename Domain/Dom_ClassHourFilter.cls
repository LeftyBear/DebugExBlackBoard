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
    ClassHourType   As Dom_ClassHourType
    Subject         As Dom_Subject
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    This.ClassHourType = ClassHourType
    Set This.Subject = Subject
    Set This.Period = Period
    Set This.Grade = Grade
End Sub

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_ClassHourFilter
    Dim Result As Dom_ClassHourFilter
    Set Result = New Dom_ClassHourFilter
    Result.InitializeInternal ClassHourType, Subject, Period, Grade
    Set Create = Result
End Function

Public Function Match(ByVal ID As Dom_ClassHourID) As Boolean
    If This.ClassHourType <> ID.ClassHourType Then Exit Function
    If Not This.Subject Is Nothing Then
        If Not This.Subject.Equals(ID.Subject) Then Exit Function
    End If
    If Not This.Period Is Nothing Then
        If Not This.Period.Equals(ID.Period) Then Exit Function
    End If
    If Not This.Grade Is Nothing Then
        If Not This.Grade.Equals(ID.Grade) Then Exit Function
    End If
    Match = True
End Function
