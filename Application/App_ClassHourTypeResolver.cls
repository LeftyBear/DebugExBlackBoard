VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourTypeResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Resolver")
Option Explicit

Public Enum App_ClassHourTypeKeyCode
    PlanKeyCode = 1
    ExeKeyCode
End Enum

Private Const KEY_PLAN      As String = "Plan"
Private Const KEY_EXECUTION As String = "Exe"

Public Function ParseKey(ByVal KeyCode As App_ClassHourTypeKeyCode) As String
    If KeyCode = PlanKeyCode Then
        ParseKey = KEY_PLAN
    ElseIf KeyCode = ExeKeyCode Then
        ParseKey = KEY_EXECUTION
    End If
End Function
