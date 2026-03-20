VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    Values() As Long
End Type

Private This As Member

Public Sub Initialize(ByVal MaxType As Long, ByVal MaxGrade As Long, ByVal MaxClassNo As Long, ByVal MaxGender As Long)
    ReDim This.Values(MaxType, MaxGrade, MaxClassNo, MaxGender)
End Sub

Public Sub AddValue(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender, ByVal Value As Long)
    This.Values(StreamType, Grade.Value, ClassNo.Value, Gender) = Value
End Sub

Public Function GradeCount() As Long
    GradeCount = UBound(This.Values, 2)
End Function

Public Function ClassNoCount() As Long
    ClassNoCount = UBound(This.Values, 3)
End Function

Public Function GetValue(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender) As Long
    GetValue = This.Values(StreamType, Grade.Value, ClassNo.Value, Gender)
End Function

Public Function UpdateValue(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender, ByVal NewValue As Long) As Long
    Dim OldValue As Long
    OldValue = This.Values(StreamType, Grade.Value, ClassNo.Value, Gender)
    This.Values(StreamType, Grade.Value, ClassNo.Value, Gender) = NewValue
    UpdateValue = NewValue - OldValue
End Function
