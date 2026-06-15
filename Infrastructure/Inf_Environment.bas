Attribute VB_Name = "Inf_Environment"
'@Folder("Infrastructure.Service")
Option Explicit
Option Private Module
'ŠJ”­’†‚Í IsDebug = True ‚Æ‚·‚é
Private Const IsDebug As Boolean = True

Public Enum Inf_EnvironmentTypeCode
    DebugMode = 1
    ReleaseMode
End Enum

Public Function GetEnvironmentTypeCode() As Inf_EnvironmentTypeCode
    If IsDebug Then
        GetEnvironmentTypeCode = DebugMode
    Else
        GetEnvironmentTypeCode = ReleaseMode
    End If
End Function
