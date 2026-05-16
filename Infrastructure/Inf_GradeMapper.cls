VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_GradeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function ToRaw(ByVal Grade As Dom_Grade) As String
    ToRaw = CStr(Grade.Value) & jpnNen
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Grade
    If 0 < VBA.InStr(1, Raw, jpnNen) Then
        Dim Grade As Long
        Grade = GetGrade(Raw)
    ElseIf 0 < VBA.InStr(1, Raw, jpnTokushi) Then
        Grade = 1
    End If
    Set ToDomain = This.Factory.CreateGrade(Grade)
End Function

Public Function ToModel(ByVal Raw As String) As Long
    ToModel = GetGrade(Raw)
End Function

Private Function GetGrade(ByVal RawGrade As String) As Long
    Dim Buffer As String
    Buffer = VBA.Replace(RawGrade, jpnNen, vbNullString)
    GetGrade = Inf_NumericUtility.ExtractNumber(Buffer)
End Function
