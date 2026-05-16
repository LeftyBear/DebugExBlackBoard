VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourIDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal ID As Dom_ClassHourID) As String
    Dim TypeMapper As New Inf_ClassHourTypeMapper
    Dim TypeCode As String
    TypeCode = TypeMapper.ToRaw(ID.TypeCode)
    Dim SubjectMapper As New Inf_SubjectMapper
    Dim Subject As String
    Subject = SubjectMapper.ToRaw(ID.Subject)
    Dim GradeMapper As New Inf_GradeMapper
    Dim Grade As String
    Grade = GradeMapper.ToRaw(ID.Grade)
    ToRaw = TypeCode & charUnderScore & Subject & charUnderScore & Grade
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourID
    Dim Parts() As String
    Parts = VBA.Split(Raw, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnKeikaku) Or 0 < VBA.InStr(1, Part, jpnJisshi) Then
            Dim TypeMapper As New Inf_ClassHourTypeMapper
            Dim ClassHourType As Dom_ClassHourType
            Set ClassHourType = TypeMapper.ToDomain(Part)
        ElseIf VBA.InStr(1, Part, jpnNen) = 0 Then
            Dim GradeMapper As New Inf_GradeMapper
            Dim Grade As Dom_Grade
            Set Grade = GradeMapper.ToDomain(Part)
        Else
            Dim SubjectMapper As New Inf_SubjectMapper
            Dim Subject As Dom_SubjectNew
            Set Subject = SubjectMapper.ToDomain(Part)
        End If
    Next
    Dim Factory As New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateClassHourID(ClassHourType, Subject, Grade)
End Function
