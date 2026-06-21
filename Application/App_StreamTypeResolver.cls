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

Private Const MODEL_MAIN    As String = "Main"
Private Const MODEL_SPECIAL As String = "Special"

Public Function ParseKey(ByVal KeyCode As Dom_StreamTypeCode) As String
    If KeyCode = MainCode Then
        ParseKey = MODEL_MAIN
    ElseIf KeyCode = SpecialCode Then
        ParseKey = MODEL_SPECIAL
    End If
End Function
