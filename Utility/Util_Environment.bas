Attribute VB_Name = "Util_Environment"
'@Folder "Utility.Environment"
Option Explicit
Option Private Module

Private Const IsDebug As Boolean = True
Public Enum Util_EnvironmentType
    DebugMode
    ReleaseMode
End Enum

Public Function GetEnvironment() As Util_EnvironmentType
    If IsDebug Then
        GetEnvironment = DebugMode
    Else
        GetEnvironment = ReleaseMode
    End If
End Function

