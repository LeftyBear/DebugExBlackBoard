Attribute VB_Name = "Inf_Environment"
'@Folder "Utility.Environment"
Option Explicit
Option Private Module

Private Const IsDebug As Boolean = True
Public Enum Util_EnvironmentType
    DebugMode
    ReleaseMode
End Enum

Public Function GetEnvironmentType() As Util_EnvironmentType
    If IsDebug Then
        GetEnvironmentType = DebugMode
    Else
        GetEnvironmentType = ReleaseMode
    End If
End Function

