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
Private Const MODEL_MAIN        As String = "Main"
Private Const MODEL_SPECIAL     As String = "Special"

Public Property Get MainKey() As String
    MainKey = MODEL_MAIN
End Property

Public Property Get SpecialKey() As String
    SpecialKey = MODEL_SPECIAL
End Property

Public Function ToRaw(ByVal TypeCode As Dom_StreamType) As String
    If TypeCode.IsMain Then
        ToRaw = vbNullString
    ElseIf TypeCode.IsSpecial Then
        ToRaw = COLUMN_SPECIAL
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_StreamType
    '@Ignore SelfAssignedDeclaration
    Dim GradeMapper As New Inf_GradeMapper
    If 0 < VBA.InStr(1, Raw, GradeMapper.ColumnName) Then
        Dim TypeCode As Long
        TypeCode = Dom_StreamTypePolicy.Main
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_SPECIAL) Then
        TypeCode = Dom_StreamTypePolicy.Special
    Else
        Exit Function
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateStreamType(TypeCode)
End Function

Public Function ToModel(ByVal Raw As String) As String
    '@Ignore SelfAssignedDeclaration
    Dim GradeMapper As New Inf_GradeMapper
    If 0 < VBA.InStr(1, Raw, GradeMapper.ColumnName) Then
        ToModel = MODEL_MAIN
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_SPECIAL) Then
        ToModel = MODEL_SPECIAL
    End If
End Function
