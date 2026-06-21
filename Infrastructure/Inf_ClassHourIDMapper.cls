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
    Dim Grade As String
    Grade = This.GradeMapper.ToRaw(ID.Grade)
    ToRaw = TypeString & DELIMITER & ID.Subject.Value & DELIMITER & Grade
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourID
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.TypeMapper.ToDomain(Raw)
    If ClassHourType Is Nothing Then Exit Function
    Dim Grade As Dom_Grade
    Set Grade = This.GradeMapper.ToDomain(Raw)
    Dim Buffer As String
    Buffer = VBA.Replace(Raw, This.TypeMapper.ToRaw(ClassHourType), vbNullString)
    Buffer = VBA.Replace(Buffer, This.GradeMapper.ToRaw(Grade), vbNullString)
    Buffer = VBA.Replace(Buffer, "_", vbNullString)
    Dim SubjectName As Dom_SubjectName
    Set SubjectName = Dom_ValueObjectFactory.CreateSubjectName(Buffer)
    Set ToDomain = Dom_ValueObjectFactory.CreateClassHourID(ClassHourType, SubjectName, Grade)
End Function
