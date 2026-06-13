VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_GradeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_GRADE As String = "”N"

Public Property Get ColumnName() As String
    ColumnName = COLUMN_GRADE
End Property

Public Function ToRaw(ByVal Grade As Dom_Grade) As String
    ToRaw = CStr(Grade.Value) & COLUMN_GRADE
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Grade
    Dim TypeMapper As Inf_StreamTypeMapper
    Set TypeMapper = New Inf_StreamTypeMapper
    Dim TypeCode As Dom_StreamType
    Set TypeCode = TypeMapper.ToDomain(Raw)
    If TypeCode Is Nothing Then Exit Function
    If TypeCode.IsMain Then
        Dim Grade As Long
        Grade = ParseGrade(Raw)
    ElseIf TypeCode.IsSpecial Then
        Grade = 1
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateGrade(Grade)
End Function

Public Function ToModel(ByVal Raw As String) As Long
    ToModel = ParseGrade(Raw)
End Function

Private Function ParseGrade(ByVal Raw As String) As Long
    Dim Buffer As String
    Buffer = VBA.Replace(Raw, COLUMN_GRADE, vbNullString)
    ParseGrade = Inf_NumericUtility.ExtractNumber(Buffer)
End Function
