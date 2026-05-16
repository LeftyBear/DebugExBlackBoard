VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_GenderMapper"
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

Public Function ToRaw(ByVal Gender As Dom_Gender) As String
    If Gender.IsMale Then
        ToRaw = jpnOtoko
    ElseIf Gender.IsFemale Then
        ToRaw = jpnOnna
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Gender
    Dim Gender As Long
    If 0 < VBA.InStr(1, Raw, jpnOtoko) Then
        Gender = Dom_GenderTypePolicy.Male
    ElseIf 0 < VBA.InStr(1, Raw, jpnOnna) Then
        Gender = Dom_GenderTypePolicy.Female
    End If
    Set ToDomain = This.Factory.CreateGender(Gender)
End Function

Public Function ToModel(ByVal Raw As String) As String
    If 0 < VBA.InStr(1, Raw, jpnOtoko) Then
        ToModel = "Male"
    ElseIf 0 < VBA.InStr(1, Raw, jpnOnna) Then
        ToModel = "Female"
    End If
End Function
