VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_GenderResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Resolver")
Option Explicit

Private Const KEY_MALE    As String = "Male"
Private Const KEY_FEMALE  As String = "Female"

Public Function ParseKey(ByVal KeyCode As Dom_GenderTypeCode) As String
    If KeyCode = MaleCode Then
        ParseKey = KEY_MALE
    ElseIf KeyCode = FemaleCode Then
        ParseKey = KEY_FEMALE
    End If
End Function
