VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourTypeModelResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Resolver")
Option Explicit

Public Enum App_ClassHourKeyCode
    PlanKeyCode = 1
    ExeKeyCode
End Enum

Private Const MODEL_PLAN        As String = "Plan"
Private Const MODEL_EXECUTION   As String = "Exe"

Public Function ToModel(ByVal ClassHourType As Dom_ClassHourType) As String
    If ClassHourType.IsPlan Then
        ToModel = MODEL_PLAN
    ElseIf ClassHourType.IsExecution Then
        ToModel = MODEL_EXECUTION
    End If
End Function

Public Function ParseKey(ByVal KeyCode As App_ClassHourKeyCode) As String
    If KeyCode = PlanKeyCode Then
        ParseKey = MODEL_PLAN
    ElseIf KeyCode = ExeKeyCode Then
        ParseKey = MODEL_EXECUTION
    End If
End Function
