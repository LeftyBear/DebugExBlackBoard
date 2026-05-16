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
    Dim Structure As App_LoadStructureUseCase
    Set Structure = CreateStructureUseCase
    Dim ClassHour As App_LoadClassHourUseCase
    Set ClassHour = CreateClassHourUseCase
    Dim Enrollment As App_LoadEnrollmentUseCase
    Set Enrollment = CreateEnrollmentUseCase(CSV)
    Dim Schedule As App_LoadScheduleUseCase
    Set Schedule = CreateScheduleUseCase
    'Presentation ----------------------------------------------------------
    Dim CalenderPresenter As Pre_CalenderPresenter
    Set CalenderPresenter = New Pre_CalenderPresenter
    CalenderPresenter.Initialize Logger, Structure, ClassHour, Enrollment, Schedule, New Pre_BasePresenter
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

Private Function CreateStructureUseCase() As App_LoadStructureUseCase
    Dim SchoolEvent As App_LoadSchoolEventUseCase
    Set SchoolEvent = CreateSchoolEventUseCase
    Dim SpecialStream As App_LoadSpecialStreamUseCase
    Set SpecialStream = CreateSpecialStreamUseCase
    Dim Subject As App_LoadSubjectUseCase
    Set Subject = CreateSubjectUseCase
    Dim UpperValue As App_LoadUpperValueUseCase
    Set UpperValue = CreateUpperValueUseCase
    Dim Factory As App_StructureUseCaseFactory
    Set Factory = New App_StructureUseCaseFactory
    Factory.Initialize SchoolEvent, SpecialStream, Subject, UpperValue
    Set CreateStructureUseCase = Factory.Create
End Function

Private Function CreateSchoolEventUseCase() As App_LoadSchoolEventUseCase
    Dim Factory As App_SchoolEventUseCaseFactory
    Set Factory = New App_SchoolEventUseCaseFactory
    Factory.Initialize New Inf_SchoolEventRepository, New App_SchoolEventRowBuilder
    Set CreateSchoolEventUseCase = Factory.Create
End Function

Private Function CreateSpecialStreamUseCase() As App_LoadSpecialStreamUseCase
    Dim Factory As App_SpecialStreamUseCaseFactory
    Set Factory = New App_SpecialStreamUseCaseFactory
    Factory.Initialize New Inf_SpecialStreamRepository, New App_SpecialStreamRowBuilder
    Set CreateSpecialStreamUseCase = Factory.Create
End Function

Private Function CreateSubjectUseCase() As App_LoadSubjectUseCase
    Dim Factory As App_SubjectUseCaseFactory
    Set Factory = New App_SubjectUseCaseFactory
    Factory.Initialize New Inf_SubjectRepository, New App_SubjectRowBuilder
    Set CreateSubjectUseCase = Factory.Create
End Function

Private Function CreateUpperValueUseCase() As App_LoadUpperValueUseCase
    Dim Factory As App_UpperValueUseCaseFactory
    Set Factory = New App_UpperValueUseCaseFactory
    Factory.Initialize New Inf_UpperValueRepository, New App_UpperValueRowBuilder
    Set CreateUpperValueUseCase = Factory.Create
End Function

Private Function CreateClassHourUseCase() As App_LoadClassHourUseCase
    Dim Factory As App_ClassHourUseCaseFactory
    Set Factory = New App_ClassHourUseCaseFactory
    Factory.Initialize New Inf_ClassHourRepository, New App_ClassHourRowBuilder
    Set CreateClassHourUseCase = Factory.Create
End Function

Private Function CreateEnrollmentUseCase(ByVal CSV As Inf_CSVPersistence) As App_LoadEnrollmentUseCase
    Dim Factory As App_EnrollmentUseCaseFactory
    Set Factory = New App_EnrollmentUseCaseFactory
    Factory.Initialize CSV
    Set CreateEnrollmentUseCase = Factory.Create
End Function

Private Function CreateScheduleUseCase() As App_LoadScheduleUseCase
    Dim Factory As App_ScheduleUseCaseFactory
    Set Factory = New App_ScheduleUseCaseFactory
    Factory.Initialize New Inf_ScheduleRepository, New App_ScheduleRowBuilder
    Set CreateScheduleUseCase = Factory.Create
End Function
