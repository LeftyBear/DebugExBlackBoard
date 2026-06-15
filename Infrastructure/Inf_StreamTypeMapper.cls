VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_StreamTypeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_SPECIAL    As String = "“ÁŽx"

Public Function ToRaw(ByVal TypeCode As Dom_StreamType) As String
    If TypeCode.IsMain Then
        ToRaw = vbNullString
    ElseIf TypeCode.IsSpecial Then
        ToRaw = COLUMN_SPECIAL
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_StreamType
    Dim TypeCode As Dom_StreamTypeCode
    Dim GradeMapper As Inf_GradeMapper
    Set GradeMapper = New Inf_GradeMapper
    If 0 < VBA.InStr(1, Raw, GradeMapper.ColumnName) Then
        TypeCode = MainCode
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_SPECIAL) Then
        TypeCode = SpecialCode
    Else
        Exit Function
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateStreamType(TypeCode)
End Function
