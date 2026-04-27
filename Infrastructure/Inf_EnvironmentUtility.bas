Attribute VB_Name = "Inf_EnvironmentUtility"
'@Folder("Infrastructure.Utility")
Option Explicit
Option Private Module
'ŠJ”­’†‚Í IsDebug = True ‚Æ‚·‚é
Private Const IsDebug As Boolean = True

Public Enum Inf_EnvironmentTypeEnum
    DebugMode = 1
    ReleaseMode
End Enum

Public Function GetEnvironmentType() As Inf_EnvironmentTypeEnum
    If IsDebug Then
        GetEnvironmentType = DebugMode
    Else
        GetEnvironmentType = ReleaseMode
    End If
End Function
