VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_StreamTypeResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Resolver")
Option Explicit

Public Enum App_StreamTypeKeyCode
    MainKeyCode = 1
    SpecialKeyCode
End Enum

Private Const MODEL_MAIN    As String = "Main"
Private Const MODEL_SPECIAL As String = "Special"

Public Function ParseKey(ByVal KeyCode As App_StreamTypeKeyCode) As String
    If KeyCode = MainKeyCode Then
        ParseKey = MODEL_MAIN
    ElseIf KeyCode = SpecialKeyCode Then
        ParseKey = MODEL_SPECIAL
    End If
End Function
