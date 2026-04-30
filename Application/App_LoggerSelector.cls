VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoggerSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Service"
Option Explicit

Public Function SelectLogger(ByVal TypeCode As Inf_EnvironmentTypePolicy, ByVal FileLogger As Inf_FileLogger, ByVal DebugLogger As Inf_DebugLogger) As App_ILogger
    If TypeCode = Inf_EnvironmentTypePolicy.ReleaseMode Then
        Set SelectLogger = FileLogger
    ElseIf TypeCode = Inf_EnvironmentTypePolicy.DebugMode Then
        Set SelectLogger = DebugLogger
    End If
End Function
