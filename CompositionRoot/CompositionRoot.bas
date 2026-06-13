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
    Dim SchedulePersis As Inf_SchedulePersistence
    Set SchedulePersis = CreateSchedulePersistence(CSV)
    Dim SchoolEventPersis As Inf_SchoolEventPersistence
    Set SchoolEventPersis = CreateSchoolEventPersistence(CSV)
    Dim ClassHourPersis As Inf_ClassHourPersistence
    Set ClassHourPersis = CreateClassHourPersistence(CSV)
    Dim SubjectPersis As Inf_SubjectPersistence
    Set SubjectPersis = CreateSubjectPersistence(CSV)
    Dim PeriodPersis As Inf_PeriodPersistence
    Set PeriodPersis = CreatePeriodPersistence(CSV)
    Dim EnrollmentPersis As Inf_EnrollmentPersistence
    Set EnrollmentPersis = CreateEnrollmentPersistence(CSV)
    Dim MainStreamPersis As Inf_MainStreamPersistence
    Set MainStreamPersis = CreateMainStreamPersistence(CSV)
    Dim SpecialStreamPersis As Inf_SpecialStreamPersistence
    Set SpecialStreamPersis = CreateSpecialStreamPersistence(CSV)
    'QueryService ----------------------------------------------------------
    Dim ScheduleQS As App_IScheduleQueryService
    Set ScheduleQS = CreateScheduleQueryService(SchedulePersis)
    Dim SchoolEventQS As App_ISchoolEventQueryService
    Set SchoolEventQS = CreateSchoolEventQueryService(SchoolEventPersis)
    Dim ClassHourQS As App_IClassHourQueryService
    Set ClassHourQS = CreateClassHourQueryService(ClassHourPersis)
    Dim SubjectQS As App_ISubjectQueryService
    Set SubjectQS = CreateSubjectQueryService(SubjectPersis)
    Dim PeriodQS As App_IPeriodQueryService
    Set PeriodQS = CreatePeriodQueryService(PeriodPersis)
    Dim EnrollmentQS As App_IEnrollmentQueryService
    Set EnrollmentQS = CreateEnrollmentQueryService(EnrollmentPersis)
    Dim MainStreamQS As App_IMainStreamQueryService
    Set MainStreamQS = CreateMainStreamQueryService(MainStreamPersis)
    Dim SpecialStreamQS As App_ISpecialStreamQueryService
    Set SpecialStreamQS = CreateSpecialStreamQueryService(SpecialStreamPersis)
    'Repository ------------------------------------------------------------
    Dim ScheduleRepo As Dom_IScheduleRepository
    Set ScheduleRepo = CreateScheduleRepository(SchedulePersis)
    Dim SchoolEventRepo As Dom_ISchoolEventRepository
    Set SchoolEventRepo = CreateSchoolEventRepository(SchoolEventPersis)
    Dim ClassHourRepo As Dom_IClassHourRepository
    Set ClassHourRepo = CreateClassHourRepository(ClassHourPersis)
    Dim SubjectRepo As Dom_ISubjectRepository
    Set SubjectRepo = CreateSubjectRepository(SubjectPersis)
    Dim PeriodRepo As Dom_IPeriodRepository
    Set PeriodRepo = CreatePeriodRepository(PeriodPersis)
    Dim EnrollmentRepo As Dom_IEnrollmentRepository
    Set EnrollmentRepo = CreateEnrollmentRepository(EnrollmentPersis)
    Dim MainStreamRepo As Dom_IMainStreamRepository
    Set MainStreamRepo = CreateMainStreamRepository(MainStreamPersis)
    Dim SpecialStreamRepo As Dom_ISpecialStreamRepository
    Set SpecialStreamRepo = CreateSpecialStreamRepository(SpecialStreamPersis)
    'Presenter -------------------------------------------------------------
    Dim MainView As Pre_MainView
    Set MainView = New Pre_MainView
    Dim DailyPeriodPre As Pre_DailyPeriodPresenter
    Set DailyPeriodPre = New Pre_DailyPeriodPresenter
    DailyPeriodPre.Inject MainView, New App_PeriodFormatter
    Dim DailySchedulePre As Pre_DailySchedulePresenter
    Set DailySchedulePre = New Pre_DailySchedulePresenter
    DailySchedulePre.Inject MainView
    'UseCaseFactory --------------------------------------------------------
    Dim Base As App_BaseUseCase
    Set Base = New App_BaseUseCase
    Dim UserUCFactory As App_UserUseCaseFactory
    Set UserUCFactory = New App_UserUseCaseFactory
    UserUCFactory.Inject Logger, ClassHourQS, ScheduleQS, MainStreamQS, Base, DailyPeriodPre, DailySchedulePre
    Dim EditerUCFactory As App_EditerUseCaseFactory
    Set EditerUCFactory = New App_EditerUseCaseFactory
    EditerUCFactory.Inject Logger, ScheduleRepo, SchoolEventRepo, ClassHourRepo, SubjectRepo, PeriodRepo, EnrollmentRepo, MainStreamRepo, SpecialStreamRepo, Base, DailyPeriodPre, DailySchedulePre
    MainView.Inject UserUCFactory
    MainView.OnChangeDate Date
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
