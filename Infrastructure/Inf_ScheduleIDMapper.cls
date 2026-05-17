VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleIDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Private Type Member
    Factory         As Dom_ValueObjectFactory
    TypeMapper      As Inf_StreamTypeMapper
    GradeMapper     As Inf_GradeMapper
    ClassNoMapper   As Inf_ClassNoMapper
    GenderMapper    As Inf_GenderMapper
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
    Set This.TypeMapper = New Inf_StreamTypeMapper
    Set This.GradeMapper = New Inf_GradeMapper
    Set This.ClassNoMapper = New Inf_ClassNoMapper
    Set This.GenderMapper = New Inf_GenderMapper
    This.TypeMapper.Initialize This.Factory
    This.GradeMapper.Initialize This.Factory
    This.ClassNoMapper.Initialize This.Factory
    This.GenderMapper.Initialize This.Factory
End Sub

Public Function ToRaw(ByVal ID As Dom_ScheduleID) As String
    Dim TypeCode As String
    TypeCode = This.TypeMapper.ToRaw(ID.StreamType)
    Dim Grade As String
    Grade = This.GradeMapper.ToRaw(ID.Grade)
    Dim ClassNo As String
    ClassNo = This.ClassNoMapper.ToRaw(ID.ClassNo)
    Dim Gender As String
    Gender = This.GenderMapper.ToRaw(ID.Gender)
    If ID.IsMain Then
        ToRaw = Grade & charUnderScore & ClassNo & charUnderScore & Gender
    ElseIf ID.IsSpecial Then
        ToRaw = TypeCode & charUnderScore & ClassNo & charUnderScore & Gender
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ScheduleID
    Dim Parts() As String
    Parts = VBA.Split(Raw, charUnderScore)
    If UBound(Parts) <> 2 Then Exit Function
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnTokushi) Or 0 < VBA.InStr(1, Part, jpnNen) Then
            Dim StreamType As Dom_StreamType
            Set StreamType = This.TypeMapper.ToDomain(Part)
            Dim Grade As Dom_Grade
            Set Grade = This.GradeMapper.ToDomain(Part)
        ElseIf 0 < VBA.InStr(1, Part, jpnKumi) Then
            Dim ClassNo As Dom_ClassNo
            Set ClassNo = This.ClassNoMapper.ToDomain(Part)
        ElseIf 0 < VBA.InStr(1, Part, jpnOtoko) Or 0 < VBA.InStr(1, Part, jpnOnna) Then
            Dim Gender As Dom_Gender
            Set Gender = This.GenderMapper.ToDomain(Part)
        End If
    Next
    Set ToDomain = This.Factory.CreateScheduleID(StreamType, Grade, ClassNo, Gender)
End Function
