Attribute VB_Name = "CompositionRoot"
'@Folder("CompositionRoot")
Option Explicit
Option Private Module

Public Sub Boot()
    'ErrorLogger -----------------------------------------------------------
    Dim LogFilePath As String
    LogFilePath = BuildLogFilePath
    If VBA.Len(VBA.Dir(LogFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "CompositionRoot", "File path is invalid.": Exit Sub
    Dim Logger As App_ILogPersistence
    Set Logger = CreateLogger(LogFilePath)
    On Error GoTo ErrHandle
    'UseCase ---------------------------------------------------------------
    Dim CSV As Inf_CSVPersistence
    Set CSV = New Inf_CSVPersistence
    'Presentation ----------------------------------------------------------
    Dim CalenderPresenter As Pre_CalenderPresenter
    Set CalenderPresenter = New Pre_CalenderPresenter
    CalenderPresenter.Initialize Logger, New Pre_BasePresenter
    Dim SchedulePresenter As Pre_SchedulePresenter
    Set SchedulePresenter = New Pre_SchedulePresenter
    SchedulePresenter.Initialize Logger, New App_ViewDTOFactory, New Pre_BasePresenter
    Dim Controller As Pre_CalenderController
    Set Controller = New Pre_CalenderController
    Controller.Initialize CalenderPresenter, SchedulePresenter
    Dim MainView As Pre_MainView
    Set MainView = New Pre_MainView
    MainView.Initialize Controller
    CalenderPresenter.AttachView MainView
    SchedulePresenter.AttachView MainView
    MainView.Show
    Exit Sub
ErrHandle:
    Logger.Log Err.Source & vbTab & Err.Description
End Sub

Private Function BuildLogFilePath() As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim BaseFolderPath As String
    BaseFolderPath = Provider.GetBaseFolderPath
    BuildLogFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "data", "errorlog", "error.log")
End Function

Private Function CreateLogger(ByVal LogFilePath As String) As App_ILogPersistence
    Dim Result As App_ILogPersistence
    Dim TypeCode As Inf_EnvironmentTypePolicy
    TypeCode = Inf_Environment.GetEnvironmentTypeCode
    If TypeCode = Inf_EnvironmentTypePolicy.ReleaseMode Then
        Dim Persistence As Inf_LogPersistence
        Set Persistence = New Inf_LogPersistence
        Persistence.Initialize LogFilePath
        Set Result = Persistence
    ElseIf TypeCode = Inf_EnvironmentTypePolicy.DebugMode Then
        Set Result = New Inf_DebugLogger
    End If
    Set CreateLogger = Result
End Function
