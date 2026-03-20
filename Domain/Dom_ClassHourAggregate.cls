VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    Hours() As Long
    Marks() As String
End Type

Private This As Member

Public Sub Initialize(ByVal MaxType As Long, ByVal MaxSubject As Long, ByVal MaxPeriod As Long, ByVal MaxGrade As Long)
    ReDim This.Hours(MaxType, MaxSubject, MaxGrade)
    ReDim This.Marks(MaxType, MaxPeriod, MaxGrade)
End Sub

Public Sub AddHours(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade, ByVal Hours As Long)
    This.Hours(ClassHourType, Subject.Index, Grade.Value) = Hours
End Sub

Public Sub AddMark(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade, ByVal Mark As String)
    This.Marks(ClassHourType, Period.Value, Grade.Value) = Mark
End Sub

Public Function GetHours(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Long
    GetHours = This.Hours(ClassHourType, Subject.Index, Grade.Value)
End Function

Public Function GetMark(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As String
    GetMark = This.Marks(ClassHourType, Period.Value, Grade.Value)
End Function

Public Function MarkCount(ByVal ClassHourType As Dom_ClassHourType, ByVal Grade As Dom_Grade) As Long
    Dim Period  As Long
    For Period = 1 To UBound(This.Marks, 2)
        If This.Marks(ClassHourType, Period, Grade.Value) <> vbNullString Then
            Dim Count As Long
            Count = Count + 1
        End If
    Next
    MarkCount = Count
End Function

Public Function SubjectCount() As Long
    SubjectCount = UBound(This.Hours, 2)
End Function

Public Function GradeCount() As Long
    GradeCount = UBound(This.Hours, 3)
End Function

Public Function UpdateHours(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade, ByVal NewValue As Long) As Long
    Dim OldValue As Long
    OldValue = This.Hours(ClassHourType, Subject.Index, Grade.Value)
    This.Hours(ClassHourType, Subject.Index, Grade.Value) = NewValue
    UpdateHours = NewValue - OldValue
End Function
