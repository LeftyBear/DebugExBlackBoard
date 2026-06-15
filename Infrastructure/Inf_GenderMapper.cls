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

Private Const COLUMN_MALE   As String = "íj"
Private Const COLUMN_FEMALE As String = "èó"

Public Function ToRaw(ByVal Gender As Dom_Gender) As String
    If Gender.IsMale Then
        ToRaw = COLUMN_MALE
    ElseIf Gender.IsFemale Then
        ToRaw = COLUMN_FEMALE
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Gender
    Dim Gender As Dom_GenderTypeCode
    If 0 < VBA.InStr(1, Raw, COLUMN_MALE) Then
        Gender = MaleCode
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_FEMALE) Then
        Gender = FemaleCode
    Else
        Exit Function
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateGender(Gender)
End Function
