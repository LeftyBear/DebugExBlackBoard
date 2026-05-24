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

Private Type Member
    Factory         As Dom_ValueObjectFactory
    TypeMapper      As Inf_ClassHourTypeMapper
    SubjectMapper   As Inf_SubjectNameMapper
    GradeMapper     As Inf_GradeMapper
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
    Set This.TypeMapper = New Inf_ClassHourTypeMapper
    Set This.SubjectMapper = New Inf_SubjectNameMapper
    Set This.GradeMapper = New Inf_GradeMapper
    This.TypeMapper.Initialize This.Factory
    This.GradeMapper.Initialize This.Factory
    This.SubjectMapper.Initialize This.Factory
End Sub

Public Function ToRaw(ByVal ID As Dom_ClassHourID) As String
    Dim TypeCode As String
    TypeCode = This.TypeMapper.ToRaw(ID.TypeCode)
    Dim SubjectName As String
    SubjectName = This.SubjectMapper.ToRaw(ID.Subject)
    Dim Grade As String
    Grade = This.GradeMapper.ToRaw(ID.Grade)
    ToRaw = TypeCode & charUnderScore & SubjectName & charUnderScore & Grade
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourID
    Dim Parts() As String
    Parts = VBA.Split(Raw, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnKeikaku) Or 0 < VBA.InStr(1, Part, jpnJisshi) Then
            Dim ClassHourType As Dom_ClassHourType
            Set ClassHourType = This.TypeMapper.ToDomain(Part)
        ElseIf VBA.InStr(1, Part, jpnNen) = 0 Then
            Dim Grade As Dom_Grade
            Set Grade = This.GradeMapper.ToDomain(Part)
        Else
            Dim SubjectName As Dom_SubjectName
            Set SubjectName = This.SubjectMapper.ToDomain(Part)
        End If
    Next
    Set ToDomain = This.Factory.CreateClassHourID(ClassHourType, SubjectName, Grade)
End Function
