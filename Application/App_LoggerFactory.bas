Attribute VB_Name = "App_LoggerFactory"
'@Folder "Application.Factory"
Option Explicit
Option Private Module

Public Function CreateLogger() As App_ILogger
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_LogFilePathBuilder
    Set Builder = New App_LogFilePathBuilder
    Builder.Initialize Provider
    Dim Resolver As App_LoggerResolver
    Set Resolver = New App_LoggerResolver
    Resolver.Initialize Builder
    Dim EnvironmentType As App_EnvironmentType
    EnvironmentType = GetEnvironmentType
    Set CreateLogger = Resolver.Resolve(EnvironmentType)
End Function
