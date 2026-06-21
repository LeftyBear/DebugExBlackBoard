VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentIDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Type Member
    TypeMapper      As Inf_StreamTypeMapper
    GradeMapper     As Inf_GradeMapper
    ClassNoMapper   As Inf_ClassNoMapper
    GenderMapper    As Inf_GenderMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.TypeMapper = New Inf_StreamTypeMapper
    Set This.GradeMapper = New Inf_GradeMapper
    Set This.ClassNoMapper = New Inf_ClassNoMapper
    Set This.GenderMapper = New Inf_GenderMapper
End Sub

Public Function ToRaw(ByVal ID As Dom_EnrollmentID) As String
    Dim Gender As String
    Gender = This.GenderMapper.ToRaw(ID.Gender)
    If ID.IsMain Then
        Dim Grade As String
        Grade = This.GradeMapper.ToRaw(ID.Grade)
        Dim ClassNo As String
        ClassNo = This.ClassNoMapper.ToRaw(ID.ClassNo)
        ToRaw = Grade & DELIMITER & ClassNo & DELIMITER & Gender
    ElseIf ID.IsSpecial Then
        Dim TypeString As String
        TypeString = This.TypeMapper.ToRaw(ID.StreamType)
        ToRaw = TypeString & DELIMITER & ID.ClassName.Value & DELIMITER & Gender
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_EnrollmentID
    Dim StreamType As Dom_StreamType
    Set StreamType = This.TypeMapper.ToDomain(Raw)
    If StreamType.IsMain Then
        Dim Grade As Dom_Grade
        Set Grade = This.GradeMapper.ToDomain(Raw)
        Dim ClassNo As Dom_ClassNo
        Set ClassNo = This.ClassNoMapper.ToDomain(Raw)
    ElseIf StreamType.IsSpecial Then
        Dim ClassName As Dom_ClassName
        Set ClassName = Dom_ValueObjectFactory.CreateClassName(Raw)
    End If
    Dim Gender As Dom_Gender
    Set Gender = This.GenderMapper.ToDomain(Raw)
    Set ToDomain = Dom_ValueObjectFactory.CreateEnrollmentID(StreamType, Grade, ClassNo, ClassName, Gender)
End Function
