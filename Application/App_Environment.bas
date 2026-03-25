Attribute VB_Name = "App_Environment"
'@Folder "Application.Model"
Option Explicit
Option Private Module

Private Const IsDebug As Boolean = True
Public Enum App_EnvironmentType
    DebugMode
    ReleaseMode
End Enum

Public Function GetEnvironmentType() As App_EnvironmentType
    If IsDebug Then
        GetEnvironmentType = DebugMode
    Else
        GetEnvironmentType = ReleaseMode
    End If
End Function
