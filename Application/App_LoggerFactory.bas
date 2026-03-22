Attribute VB_Name = "App_LoggerFactory"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Function CreateLogger() As Inf_ILogger
    Dim Builder As Inf_LogFilePathBuilder
    Set Builder = New Inf_LogFilePathBuilder
    Builder.Initialize New Inf_WorkbookPathProvider
    Dim Resolver As App_LoggerResolver
    Set Resolver = New App_LoggerResolver
    Resolver.Initialize Builder
    Dim EnvironmentType As Util_EnvironmentType
    EnvironmentType = GetEnvironmentType
    Set CreateLogger = Resolver.Resolve(EnvironmentType)
End Function
