VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoggerSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Selector"
Option Explicit

Public Function SelectLogger(ByVal EnvironmentType As Inf_EnvironmentTypeEnum, ByVal FileLogger As Inf_FileLogger, ByVal DebugLogger As Inf_DebugLogger) As App_ILogger
    Select Case EnvironmentType
    Case Inf_EnvironmentTypeEnum.ReleaseMode
        Set SelectLogger = FileLogger
    Case Else
        Set SelectLogger = DebugLogger
    End Select
End Function
