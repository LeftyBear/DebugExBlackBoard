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
    'Persistence -----------------------------------------------------------
    Dim CSV As Inf_CSVPersistence
    Set CSV = New Inf_CSVPersistence
    Dim SchedulePersistence As Inf_SchedulePersistence
    Set SchedulePersistence = CreateSchedulePersistence(CSV)
    Dim SchoolEventPersistence As Inf_SchoolEventPersistence
    Set SchoolEventPersistence = CreateSchoolEventPersistence(CSV)
    Dim ClassHourPersistence As Inf_ClassHourPersistence
    Set ClassHourPersistence = CreateClassHourPersistence(CSV)
    Dim SubjectPersistence As Inf_SubjectPersistence
    Set SubjectPersistence = CreateSubjectPersistence(CSV)
    Dim PeriodPersistence As Inf_PeriodPersistence
    Set PeriodPersistence = CreatePeriodPersistence(CSV)
    Dim EnrollmentPersistence As Inf_EnrollmentPersistence
    Set EnrollmentPersistence = CreateEnrollmentPersistence(CSV)
    Dim MainStreamPersistence As Inf_MainStreamPersistence
    Set MainStreamPersistence = CreateMainStreamPersistence(CSV)
    Dim SpecialStreamPersistence As Inf_SpecialStreamPersistence
    Set SpecialStreamPersistence = CreateSpecialStreamPersistence(CSV)
    'QueryService ----------------------------------------------------------
    Dim ScheduleQueryService As App_IScheduleQueryService
    Set ScheduleQueryService = CreateScheduleQueryService(SchedulePersistence)
    Dim SchoolEventQueryService As App_ISchoolEventQueryService
    Set SchoolEventQueryService = CreateSchoolEventQueryService(SchoolEventPersistence)
    Dim ClassHourQueryService As App_IClassHourQueryService
    Set ClassHourQueryService = CreateClassHourQueryService(ClassHourPersistence)
    Dim SubjectQueryService As App_ISubjectQueryService
    Set SubjectQueryService = CreateSubjectQueryService(SubjectPersistence)
    Dim PeriodQueryService As App_IPeriodQueryService
    Set PeriodQueryService = CreatePeriodQueryService(PeriodPersistence)
    Dim EnrollmentQueryService As App_IEnrollmentQueryService
    Set EnrollmentQueryService = CreateEnrollmentQueryService(EnrollmentPersistence)
    Dim MainStreamQueryService As App_IMainStreamQueryService
    Set MainStreamQueryService = CreateMainStreamQueryService(MainStreamPersistence)
    Dim SpecialStreamQueryService As App_ISpecialStreamQueryService
    Set SpecialStreamQueryService = CreateSpecialStreamQueryService(SpecialStreamPersistence)
    'Repository ------------------------------------------------------------
    Dim ScheduleRepository As Dom_IScheduleRepository
    Set ScheduleRepository = CreateScheduleRepository(SchedulePersistence)
    Dim SchoolEventRepository As Dom_ISchoolEventRepository
    Set SchoolEventRepository = CreateSchoolEventRepository(SchoolEventPersistence)
    Dim ClassHourRepository As Dom_IClassHourRepository
    Set ClassHourRepository = CreateClassHourRepository(ClassHourPersistence)
    Dim SubjectRepository As Dom_ISubjectRepository
    Set SubjectRepository = CreateSubjectRepository(SubjectPersistence)
    Dim PeriodRepository As Dom_IPeriodRepository
    Set PeriodRepository = CreatePeriodRepository(PeriodPersistence)
    Dim EnrollmentRepository As Dom_IEnrollmentRepository
    Set EnrollmentRepository = CreateEnrollmentRepository(EnrollmentPersistence)
    Dim MainStreamRepository As Dom_IMainStreamRepository
    Set MainStreamRepository = CreateMainStreamRepository(MainStreamPersistence)
    Dim SpecialStreamRepository As Dom_ISpecialStreamRepository
    Set SpecialStreamRepository = CreateSpecialStreamRepository(SpecialStreamPersistence)
    'UseCase ---------------------------------------------------------------
    Dim LoadDailyScheduleUseCase As App_LoadDailyScheduleUseCase
    Set LoadDailyScheduleUseCase = CreateLoadDailyScheduleUseCase(ScheduleQueryService)
    Dim TotalDailyPeriodUseCase As App_BuildPeriodTotalUseCase
    Set TotalDailyPeriodUseCase = CreateTotalDailyPeriodUseCase(ClassHourQueryService, MainStreamQueryService)
    'Presentation ----------------------------------------------------------
    Dim DailySchedulePresenter As Pre_DailySchedulePresenter
    Set DailySchedulePresenter = CreateDailySchedulePresenter(LoadDailyScheduleUseCase, Logger)
    Dim DailyPeriodPresenter As Pre_DailyPeriodPresenter
    Set DailyPeriodPresenter = CreateDailyPeriodPresenter(TotalDailyPeriodUseCase, Logger)
    'View ------------------------------------------------------------------
    Dim MainView As Pre_MainView
    Set MainView = New Pre_MainView
    MainView.Inject DailySchedulePresenter, DailyPeriodPresenter
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
    Else
        Set Result = New Inf_DebugLogger
    End If
    Set CreateLogger = Result
End Function

Private Function CreateSchedulePersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_SchedulePersistence
    Dim Result As Inf_SchedulePersistence
    Set Result = New Inf_SchedulePersistence
    Result.Inject Persistence
    Set CreateSchedulePersistence = Result
End Function

Private Function CreateScheduleQueryService(ByVal Persistence As Inf_SchedulePersistence) As App_IScheduleQueryService
    Dim Result As Inf_ScheduleQueryService
    Set Result = New Inf_ScheduleQueryService
    Result.Inject Persistence
    Set CreateScheduleQueryService = Result
End Function

Private Function CreateScheduleRepository(ByVal Persistence As Inf_SchedulePersistence) As Dom_IScheduleRepository
    Dim Result As Inf_ScheduleRepository
    Set Result = New Inf_ScheduleRepository
    Result.Inject Persistence
    Set CreateScheduleRepository = Result
End Function

Private Function CreateSchoolEventPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_SchoolEventPersistence
    Dim Result As Inf_SchoolEventPersistence
    Set Result = New Inf_SchoolEventPersistence
    Result.Inject Persistence
    Set CreateSchoolEventPersistence = Result
End Function

Private Function CreateSchoolEventQueryService(ByVal Persistence As Inf_SchoolEventPersistence) As App_ISchoolEventQueryService
    Dim Result As Inf_SchoolEventQueryService
    Set Result = New Inf_SchoolEventQueryService
    Result.Inject Persistence
    Set CreateSchoolEventQueryService = Result
End Function

Private Function CreateSchoolEventRepository(ByVal Persistence As Inf_SchoolEventPersistence) As Dom_ISchoolEventRepository
    Dim Result As Inf_SchoolEventRepository
    Set Result = New Inf_SchoolEventRepository
    Result.Inject Persistence
    Set CreateSchoolEventRepository = Result
End Function

Private Function CreateClassHourPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_ClassHourPersistence
    Dim Result As Inf_ClassHourPersistence
    Set Result = New Inf_ClassHourPersistence
    Result.Inject Persistence
    Set CreateClassHourPersistence = Result
End Function

Private Function CreateClassHourQueryService(ByVal Persistence As Inf_ClassHourPersistence) As App_IClassHourQueryService
    Dim Result As Inf_ClassHourQueryService
    Set Result = New Inf_ClassHourQueryService
    Result.Inject Persistence
    Set CreateClassHourQueryService = Result
End Function

Private Function CreateClassHourRepository(ByVal Persistence As Inf_ClassHourPersistence) As Dom_IClassHourRepository
    Dim Result As Inf_ClassHourRepository
    Set Result = New Inf_ClassHourRepository
    Result.Inject Persistence
    Set CreateClassHourRepository = Result
End Function

Private Function CreateSubjectPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_SubjectPersistence
    Dim Result As Inf_SubjectPersistence
    Set Result = New Inf_SubjectPersistence
    Result.Inject Persistence
    Set CreateSubjectPersistence = Result
End Function

Private Function CreateSubjectQueryService(ByVal Persistence As Inf_SubjectPersistence) As App_ISubjectQueryService
    Dim Result As Inf_SubjectQueryService
    Set Result = New Inf_SubjectQueryService
    Result.Inject Persistence
    Set CreateSubjectQueryService = Result
End Function

Private Function CreateSubjectRepository(ByVal Persistence As Inf_SubjectPersistence) As Dom_ISubjectRepository
    Dim Result As Inf_SubjectRepository
    Set Result = New Inf_SubjectRepository
    Result.Inject Persistence
    Set CreateSubjectRepository = Result
End Function

Private Function CreatePeriodPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_PeriodPersistence
    Dim Result As Inf_PeriodPersistence
    Set Result = New Inf_PeriodPersistence
    Result.Inject Persistence
    Set CreatePeriodPersistence = Result
End Function

Private Function CreatePeriodQueryService(ByVal Persistence As Inf_PeriodPersistence) As App_IPeriodQueryService
    Dim Result As Inf_PeriodQueryService
    Set Result = New Inf_PeriodQueryService
    Result.Inject Persistence
    Set CreatePeriodQueryService = Result
End Function

Private Function CreatePeriodRepository(ByVal Persistence As Inf_PeriodPersistence) As Dom_IPeriodRepository
    Dim Result As Inf_PeriodRepository
    Set Result = New Inf_PeriodRepository
    Result.Inject Persistence
    Set CreatePeriodRepository = Result
End Function

Private Function CreateEnrollmentPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_EnrollmentPersistence
    Dim Result As Inf_EnrollmentPersistence
    Set Result = New Inf_EnrollmentPersistence
    Result.Inject Persistence
    Set CreateEnrollmentPersistence = Result
End Function

Private Function CreateEnrollmentQueryService(ByVal Persistence As Inf_EnrollmentPersistence) As App_IEnrollmentQueryService
    Dim Result As Inf_EnrollmentQueryService
    Set Result = New Inf_EnrollmentQueryService
    Result.Inject Persistence
    Set CreateEnrollmentQueryService = Result
End Function

Private Function CreateEnrollmentRepository(ByVal Persistence As Inf_EnrollmentPersistence) As Dom_IEnrollmentRepository
    Dim Result As Inf_EnrollmentRepository
    Set Result = New Inf_EnrollmentRepository
    Result.Inject Persistence
    Set CreateEnrollmentRepository = Result
End Function

Private Function CreateMainStreamPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_MainStreamPersistence
    Dim Result As Inf_MainStreamPersistence
    Set Result = New Inf_MainStreamPersistence
    Result.Inject Persistence
    Set CreateMainStreamPersistence = Result
End Function

Private Function CreateMainStreamQueryService(ByVal Persistence As Inf_MainStreamPersistence) As App_IMainStreamQueryService
    Dim Result As Inf_MainStreamQueryService
    Set Result = New Inf_MainStreamQueryService
    Result.Inject Persistence
    Set CreateMainStreamQueryService = Result
End Function

Private Function CreateMainStreamRepository(ByVal Persistence As Inf_MainStreamPersistence) As Dom_IMainStreamRepository
    Dim Result As Inf_MainStreamRepository
    Set Result = New Inf_MainStreamRepository
    Result.Inject Persistence
    Set CreateMainStreamRepository = Result
End Function

Private Function CreateSpecialStreamPersistence(ByVal Persistence As Inf_CSVPersistence) As Inf_SpecialStreamPersistence
    Dim Result As Inf_SpecialStreamPersistence
    Set Result = New Inf_SpecialStreamPersistence
    Result.Inject Persistence
    Set CreateSpecialStreamPersistence = Result
End Function

Private Function CreateSpecialStreamQueryService(ByVal Persistence As Inf_SpecialStreamPersistence) As App_ISpecialStreamQueryService
    Dim Result As Inf_SpecialStreamQueryService
    Set Result = New Inf_SpecialStreamQueryService
    Result.Inject Persistence
    Set CreateSpecialStreamQueryService = Result
End Function

Private Function CreateSpecialStreamRepository(ByVal Persistence As Inf_SpecialStreamPersistence) As Dom_ISpecialStreamRepository
    Dim Result As Inf_SpecialStreamRepository
    Set Result = New Inf_SpecialStreamRepository
    Result.Inject Persistence
    Set CreateSpecialStreamRepository = Result
End Function

Private Function CreateLoadDailyScheduleUseCase(ByVal QueryService As Inf_ScheduleQueryService) As App_LoadDailyScheduleUseCase
    Dim Result As App_LoadDailyScheduleUseCase
    Set Result = New App_LoadDailyScheduleUseCase
    Result.Inject QueryService
    Set CreateLoadDailyScheduleUseCase = Result
End Function

Private Function CreateTotalDailyPeriodUseCase(ByVal ClassHourQS As Inf_ClassHourQueryService, ByVal MainStreamQS As App_IMainStreamQueryService) As App_BuildPeriodTotalUseCase
    Dim Result As App_BuildPeriodTotalUseCase
    Set Result = New App_BuildPeriodTotalUseCase
    Result.Inject ClassHourQS, MainStreamQS
    Set CreateTotalDailyPeriodUseCase = Result
End Function

Private Function CreateDailySchedulePresenter(ByVal UseCase As App_LoadDailyScheduleUseCase, ByVal Logger As App_ILogPersistence) As Pre_DailySchedulePresenter
    Dim Result As Pre_DailySchedulePresenter
    Set Result = New Pre_DailySchedulePresenter
    Result.Inject UseCase, Logger, New Pre_BasePresenter
    Set CreateDailySchedulePresenter = Result
End Function

Private Function CreateDailyPeriodPresenter(ByVal UseCase As App_BuildPeriodTotalUseCase, ByVal Logger As App_ILogPersistence) As Pre_DailyPeriodPresenter
    Dim Result As Pre_DailyPeriodPresenter
    Set Result = New Pre_DailyPeriodPresenter
    Result.Inject UseCase, Logger
    Set CreateDailyPeriodPresenter = Result
End Function
