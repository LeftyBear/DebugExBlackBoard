VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourIDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Type Member
    TypeMapper  As Inf_ClassHourTypeMapper
    GradeMapper As Inf_GradeMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.TypeMapper = New Inf_ClassHourTypeMapper
    Set This.GradeMapper = New Inf_GradeMapper
End Sub

Public Function ToRaw(ByVal ID As Dom_ClassHourID) As String
    Dim TypeString As String
    TypeString = This.TypeMapper.ToRaw(ID.ClassHourType)
    Dim SubjectName As String
    SubjectName = ID.Subject.Value
    Dim Grade As String
    Grade = This.GradeMapper.ToRaw(ID.Grade)
    ToRaw = TypeString & DELIMITER & SubjectName & DELIMITER & Grade
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourID
    Dim Parts() As String
    Parts = VBA.Split(Raw, DELIMITER)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        Dim ClassHourType As Dom_ClassHourType
        Dim Grade As Dom_Grade
        Dim SubjectName As Dom_SubjectName
        If ClassHourType Is Nothing Then
            Set ClassHourType = This.TypeMapper.ToDomain(Part)
        ElseIf Grade Is Nothing Then
            Set Grade = This.GradeMapper.ToDomain(Part)
        ElseIf SubjectName Is Nothing Then
            Set SubjectName = Dom_ValueObjectFactory.CreateSubjectName(Part)
        End If
    Next
    Set ToDomain = Dom_ClassHourIDFactory.Create(ClassHourType, SubjectName, Grade)
End Function
