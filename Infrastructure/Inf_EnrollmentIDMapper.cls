VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentIDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal ID As Dom_EnrollmentID) As String
    Dim TypeMapper As New Inf_StreamTypeMapper
    Dim TypeCode As String
    TypeCode = TypeMapper.ToRaw(ID.StreamType)
    Dim GradeMapper As New Inf_GradeMapper
    Dim Grade As String
    Grade = GradeMapper.ToRaw(ID.Grade)
    Dim ClassNoMapper As New Inf_ClassNoMapper
    Dim ClassNo As String
    ClassNo = ClassNoMapper.ToRaw(ID.ClassNo)
    Dim GenderMapper As New Inf_GenderMapper
    Dim Gender As String
    Gender = GenderMapper.ToRaw(ID.Gender)
    If ID.IsMain Then
        ToRaw = Grade & charUnderScore & ClassNo & charUnderScore & Gender
    ElseIf ID.IsSpecial Then
        ToRaw = TypeCode & charUnderScore & ClassNo & charUnderScore & Gender
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_EnrollmentID
    Dim Parts() As String
    Parts = VBA.Split(Raw, charUnderScore)
    If UBound(Parts) <> 2 Then Exit Function
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnTokushi) Or 0 < VBA.InStr(1, Part, jpnNen) Then
            Dim TypeMapper As New Inf_StreamTypeMapper
            Dim StreamType As Dom_StreamType
            Set StreamType = TypeMapper.ToDomain(Part)
            Dim GradeMapper As New Inf_GradeMapper
            Dim Grade As Dom_Grade
            Set Grade = GradeMapper.ToDomain(Part)
        ElseIf 0 < VBA.InStr(1, Part, jpnKumi) Then
            Dim ClassNoMapper As New Inf_ClassNoMapper
            Dim ClassNo As Dom_ClassNo
            Set ClassNo = ClassNoMapper.ToDomain(Part)
        ElseIf 0 < VBA.InStr(1, Part, jpnOtoko) Or 0 < VBA.InStr(1, Part, jpnOnna) Then
            Dim GenderMapper As New Inf_GenderMapper
            Dim Gender As Dom_Gender
            Set Gender = GenderMapper.ToDomain(Part)
        End If
    Next
    Dim Factory As New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateEnrollmentID(StreamType, Grade, ClassNo, Gender)
End Function
