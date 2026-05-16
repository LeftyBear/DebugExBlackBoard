Attribute VB_Name = "Inf_Environment"
'@Folder("Infrastructure.Service")
Option Explicit
Option Private Module
'ŠJ”­’†‚Í IsDebug = True ‚Æ‚·‚é
Private Const IsDebug As Boolean = True

Public Enum Inf_EnvironmentTypePolicy
    DebugMode = 1
    ReleaseMode
End Enum

Public Function GetEnvironmentTypeCode() As Inf_EnvironmentTypePolicy
    If IsDebug Then
        GetEnvironmentTypeCode = Inf_EnvironmentTypePolicy.DebugMode
    Else
        GetEnvironmentTypeCode = Inf_EnvironmentTypePolicy.ReleaseMode
    End If
End Function
