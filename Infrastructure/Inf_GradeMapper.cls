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

Public Function ToRaw(ByVal Grade As Dom_Grade) As String
    ToRaw = CStr(Grade.Value) & jpnNen
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Grade
    Dim Grade As Long
    If 0 < VBA.InStr(1, Raw, jpnNen) Then
        Dim Buffer As String
        Buffer = VBA.Replace(Raw, jpnNen, vbNullString)
        Grade = Inf_NumericUtility.ExtractNumber(Buffer)
    ElseIf 0 < VBA.InStr(1, Raw, jpnTokushi) Then
        Grade = 1
    End If
    Dim Factory As Dom_ValueObjectFactory
    Set Factory = New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateGrade(Grade)
End Function
