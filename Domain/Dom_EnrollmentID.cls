VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    StreamType  As Dom_StreamType
    Grade       As Dom_Grade
    ClassNo     As Dom_ClassNo
    Gender      As Dom_Gender
End Type

Private This As Member

Public Sub Initialize(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender)
    Set This.StreamType = StreamType
    Set This.Grade = Grade
    Set This.ClassNo = ClassNo
    Set This.Gender = Gender
End Sub

Public Property Get StreamType() As Dom_StreamType
    Set StreamType = This.StreamType
End Property

Public Property Get ClassNo() As Dom_ClassNo
    Set ClassNo = This.ClassNo
End Property

Public Property Get Gender() As Dom_Gender
    Set Gender = This.Gender
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Public Function ToKey() As String
    If This.StreamType.Value = Dom_StreamTypePolicy.Main Then
        ToKey = This.Grade.ToKey & charUnderScore & This.ClassNo.ToKey & charUnderScore & This.Gender.ToKey
    ElseIf This.StreamType.Value = Dom_StreamTypePolicy.Special Then
        ToKey = jpnTokushi & charUnderScore & This.ClassNo.ToKey & charUnderScore & This.Gender.ToKey
    End If
End Function

Public Function Equals(ByVal Other As Dom_EnrollmentID) As Boolean
    If Other Is Nothing Then Exit Function
    If Not This.StreamType.Equals(Other.StreamType) Then Exit Function
    If Not This.Grade.Equals(Other.Grade) Then Exit Function
    If Not This.ClassNo.Equals(Other.ClassNo) Then Exit Function
    If Not This.Gender.Equals(Other.Gender) Then Exit Function
    Equals = True
End Function

Public Function IsMainStream() As Boolean
    IsMainStream = (This.StreamType.Value = Dom_StreamTypePolicy.Main)
End Function

Public Function IsSpecialStream() As Boolean
    IsSpecialStream = (This.StreamType.Value = Dom_StreamTypePolicy.Special)
End Function
