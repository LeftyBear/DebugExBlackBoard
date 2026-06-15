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

Public Enum App_GenderKeyCode
    Male = 1
    Female
End Enum

Private Const MODEL_MALE    As String = "Male"
Private Const MODEL_FEMALE  As String = "Female"

Public Function ParseKey(ByVal KeyCode As App_GenderKeyCode) As String
    If KeyCode = Male Then
        ParseKey = MODEL_MALE
    ElseIf KeyCode = Female Then
        ParseKey = MODEL_FEMALE
    End If
End Function
