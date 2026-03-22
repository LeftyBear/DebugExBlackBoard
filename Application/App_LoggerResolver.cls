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

Private Type Member
    Builder As Inf_LogFilePathBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As Inf_LogFilePathBuilder)
    Set This.Builder = Builder
End Sub

Public Function Resolve(ByVal Environment As Util_EnvironmentType) As Inf_ILogger
    Select Case Environment
    Case Util_EnvironmentType.ReleaseMode
        Dim FilePath As String
        FilePath = This.Builder.Builed
        Dim Logger As Inf_FileLogger
        Set Logger = New Inf_FileLogger
        Logger.Initialize FilePath
        Set Resolve = Logger
    Case Else
        Set Resolve = New Inf_DebugLogger
    End Select
End Function
