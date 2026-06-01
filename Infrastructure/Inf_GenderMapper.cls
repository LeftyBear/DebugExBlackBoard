VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_GenderMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_MALE       As String = "íj"
Private Const COLUMN_FEMALE     As String = "èó"
Private Const MODEL_MALE        As String = "Male"
Private Const MODEL_FEMALE      As String = "Female"

Public Property Get MaleKey() As String
    MaleKey = MODEL_MALE
End Property

Public Property Get FemaleKey() As String
    FemaleKey = MODEL_FEMALE
End Property

Public Function ToRaw(ByVal Gender As Dom_Gender) As String
    If Gender.IsMale Then
        ToRaw = COLUMN_MALE
    ElseIf Gender.IsFemale Then
        ToRaw = COLUMN_FEMALE
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Gender
    If 0 < VBA.InStr(1, Raw, COLUMN_MALE) Then
        Dim Gender As Long
        Gender = Dom_GenderTypePolicy.Male
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_FEMALE) Then
        Gender = Dom_GenderTypePolicy.Female
    Else
        Exit Function
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateGender(Gender)
End Function

Public Function ToModel(ByVal Raw As String) As String
    If 0 < VBA.InStr(1, Raw, COLUMN_MALE) Then
        ToModel = MODEL_MALE
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_FEMALE) Then
        ToModel = MODEL_FEMALE
    End If
End Function
