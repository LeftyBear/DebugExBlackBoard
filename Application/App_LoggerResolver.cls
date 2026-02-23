VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoggerResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Resolver"
Option Explicit

Public Function Resolve(ByVal Environment As Util_EnvironmentType, ByVal LogPath As String) As Inf_ILogger
    Select Case Environment
    Case Util_EnvironmentType.ReleaseMode
        Dim Logger As Inf_FileLogger
        Set Logger = New Inf_FileLogger
        Logger.Initialize LogPath
        Set Resolve = Logger
    Case Else
        Set Resolve = New Inf_DebugLogger
    End Select
End Function

