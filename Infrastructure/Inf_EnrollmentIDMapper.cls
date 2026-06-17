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
    Dim ClassNo As String
    ClassNo = This.ClassNoMapper.ToRaw(ID.ClassNo)
    Dim Gender As String
    Gender = This.GenderMapper.ToRaw(ID.Gender)
    If ID.IsMain Then
        Dim Grade As String
        Grade = This.GradeMapper.ToRaw(ID.Grade)
        ToRaw = Grade & DELIMITER & ClassNo & DELIMITER & Gender
    ElseIf ID.IsSpecial Then
        Dim TypeString As String
        TypeString = This.TypeMapper.ToRaw(ID.StreamType)
        ToRaw = TypeString & DELIMITER & ID.ClassName & DELIMITER & Gender
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_EnrollmentID
    Dim Parts() As String
    Parts = VBA.Split(Raw, DELIMITER)
    If UBound(Parts) <> 2 Then Exit Function
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        Dim StreamType As Dom_StreamType
        If StreamType Is Nothing Then
            Set StreamType = This.TypeMapper.ToDomain(Part)
        End If
        Dim Grade As Dom_Grade
        If Grade Is Nothing Then
            Set Grade = This.GradeMapper.ToDomain(Part)
        End If
        Dim ClassNo As Dom_ClassNo
        If ClassNo Is Nothing Then
            Set ClassNo = This.ClassNoMapper.ToDomain(Part)
        End If
        Dim Gender As Dom_Gender
        If Gender Is Nothing Then
            Set Gender = This.GenderMapper.ToDomain(Part)
        End If
    Next
    Set ToDomain = Dom_EnrollmentIDFactory.Create(StreamType, Grade, ClassNo, ClassName, Gender)
End Function
