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
    Dim CSV As Inf_IPersistence
    Set CSV = New Inf_CSVPersistence
    Dim ClassHourPersistence As Inf_IClassHourPersistence
    Set ClassHourPersistence =CreateClassHourPersistence(CSV)
    Dim EnrollmentPersistence As Inf_IEnrollmentPersistence
    Set EnrollmentPersistence = CreateEnrollmentPersistence(CSV)
    Dim SchedulePersistence As Inf_ISchedulePersistence
    Set SchedulePersistence = CreateSchedulePersistence(CSV)
    Dim MainStreamPersistence As Inf_IMainStreamPersistence
    Set MainStreamPersistence = CreateMainStreamPersistence(CSV)
    Dim SpecialStreamPersistence As Inf_ISpecialStreamPersistence
    Set SpecialStreamPersistence = CreateSpecialStreamPersistence(CSV)
    Dim SubjectPersistence As Inf_ISubjectPersistence
    Set SubjectPersistence = CreateSubjectPersistence(CSV)
    Dim PeriodPersistence As Inf_IPeriodPersistence
    Set PeriodPersistence = CreatePeriodPersistence(CSV)
    Dim SchoolEventPersistence As Inf_ISchoolEventPersistence
    Set SchoolEventPersistence = CreateSchoolEventPersistence(CSV)
    'QueryService ----------------------------------------------------------
    Dim ClassHourQueryService As Inf_IClassHourQueryService
    Set ClassHourQueryService =CreateClassHourQueryService(ClassHourPersistence)
    Dim EnrollmentQueryService As Inf_IEnrollmentQueryService
    Set EnrollmentQueryService = CreateEnrollmentQueryService(EnrollmentPersistence)
    Dim ScheduleQueryService As Inf_IScheduleQueryService
    Set ScheduleQueryService = CreateScheduleQueryService(SchedulePersistence)
    Dim MainStreamQueryService As Inf_IMainStreamQueryService
    Set MainStreamQueryService = CreateMainStreamQueryService(MainStreamPersistence)
    Dim SpecialStreamQueryService As Inf_ISpecialStreamQueryService
    Set SpecialStreamQueryService = CreateSpecialStreamQueryService(SpecialStreamPersistence)
    Dim SubjectQueryService As Inf_ISubjectQueryService
    Set SubjectQueryService = CreateSubjectQueryService(SubjectPersistence)
    Dim PeriodQueryService As Inf_IPeriodQueryService
    Set PeriodQueryService = CreatePeriodQueryService(PeriodPersistence)
    Dim SchoolEventQueryService As Inf_ISchoolEventQueryService
    Set SchoolEventQueryService = CreateSchoolEventQueryService(SchoolEventPersistence)
    'Repository ------------------------------------------------------------
    Dim ClassHourRepository As Inf_IClassHourRepository
    Set ClassHourRepository = CreateClassHourRepository(ClassHourPersistence)
    Dim EnrollmentRepository As Inf_IEnrollmentRepository
    Set EnrollmentRepository = CreateEnrollmentRepository(EnrollmentPersistence)
    Dim ScheduleRepository As Inf_IScheduleRepository
    Set ScheduleRepository = CreateScheduleRepository(SchedulePersistence)
    Dim MainStreamRepository As Inf_IMainStreamRepository
    Set MainStreamRepository = CreateMainStreamRepository(MainStreamPersistence)
    Dim SpecialStreamRepository As Inf_ISpecialStreamRepository
    Set SpecialStreamRepository = CreateSpecialStreamRepository(SpecialStreamPersistence)
    Dim SubjectRepository As Inf_ISubjectRepository
    Set SubjectRepository = CreateSubjectRepository(SubjectPersistence)
    Dim PeriodRepository As Inf_IPeriodRepository
    Set PeriodRepository = CreatePeriodRepository(PeriodPersistence)
    Dim SchoolEventRepository As Inf_ISchoolEventRepository
    Set SchoolEventRepository = CreateSchoolEventRepository(SchoolEventPersistence)
    'UseCase ---------------------------------------------------------------
    Dim TotalDailyPeriodUseCase As App_TotalDailyPeriodUseCase
    Set TotalDailyPeriodUseCase = New App_TotalDailyPeriodUseCase

    Dim LoadDailyScheduleUseCase As App_LoadDailyScheduleUseCase
    Set LoadDailyScheduleUseCase = New App_LoadDailyScheduleUseCase

    'Presentation ----------------------------------------------------------

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

Private Function CreateClassHourPersistence(Byval CSV as Inf_CSVPersistence) As Inf_IClassHourParsistence
    Dim Result As Inf_IClassHourParsistence
    Set Result = New Inf_IClassHourParsistence
    Result.Inject CSV
    Set CreateClassHourPersistence = Result
End Function

Private Function CreateClassHourQueryService(ByVal ClassHourPersistence As Inf_IClassHourPersistence) As Inf_IClassHourQueryService
    Dim Result As Inf_IClassHourQueryService
    Set Result = New Inf_ClassHourQueryService
    Result.Inject ClassHourPersistence
    Set CreateClassHourQueryService = Result
End Function

Private Function CreateClassHourRepository(ByVal ClassHourPersistence As Inf_IClassHourPersistence) As Inf_IClassHourRepository
    Dim Result As Inf_IClassHourRepository
    Set Result = New Inf_ClassHourRepository
    Result.Inject ClassHourPersistence
    Set CreateClassHourRepository = Result
End Function

Private Function CreateEnrollmentPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_IEnrollmentPersistence
    Dim Result As Inf_IEnrollmentPersistence
    Set Result = New Inf_EnrollmentPersistence
    Result.Inject CSV
    Set CreateEnrollmentPersistence = Result
End Function

Private Function CreateEnrollmentQueryService(ByVal EnrollmentPersistence As Inf_IEnrollmentPersistence) As Inf_IEnrollmentQueryService
    Dim Result As Inf_IEnrollmentQueryService
    Set Result = New Inf_EnrollmentQueryService
    Result.Inject EnrollmentPersistence
    Set CreateEnrollmentQueryService = Result
End Function

Private Function CreateEnrollmentRepository(ByVal EnrollmentPersistence As Inf_IEnrollmentPersistence) As Inf_IEnrollmentRepository
    Dim Result As Inf_IEnrollmentRepository
    Set Result = New Inf_EnrollmentRepository
    Result.Inject EnrollmentPersistence
    Set CreateEnrollmentRepository = Result
End Function

Private Function CreateSchedulePersistence(ByVal CSV As Inf_CSVPersistence) As Inf_ISchedulePersistence
    Dim Result As Inf_ISchedulePersistence
    Set Result = New Inf_SchedulePersistence
    Result.Inject CSV
    Set CreateSchedulePersistence = Result
End Function

Private Function CreateScheduleQueryService(ByVal SchedulePersistence As Inf_ISchedulePersistence) As Inf_IScheduleQueryService
    Dim Result As Inf_IScheduleQueryService
    Set Result = New Inf_ScheduleQueryService
    Result.Inject SchedulePersistence
    Set CreateScheduleQueryService = Result
End Function

Private Function CreateScheduleRepository(ByVal SchedulePersistence As Inf_ISchedulePersistence) As Inf_IScheduleRepository
    Dim Result As Inf_IScheduleRepository
    Set Result = New Inf_ScheduleRepository
    Result.Inject SchedulePersistence
    Set CreateScheduleRepository = Result
End Function

Private Function CreateMainStreamPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_IMainStreamPersistence
    Dim Result As Inf_IMainStreamPersistence
    Set Result = New Inf_MainStreamPersistence
    Result.Inject CSV
    Set CreateMainStreamPersistence = Result
End Function

Private Function CreateMainStreamQueryService(ByVal MainStreamPersistence As Inf_IMainStreamPersistence) As Inf_IMainStreamQueryService
    Dim Result As Inf_IMainStreamQueryService
    Set Result = New Inf_MainStreamQueryService
    Result.Inject MainStreamPersistence
    Set CreateMainStreamQueryService = Result
End Function

Private Function CreateMainStreamRepository(ByVal MainStreamPersistence As Inf_IMainStreamPersistence) As Inf_IMainStreamRepository
    Dim Result As Inf_IMainStreamRepository
    Set Result = New Inf_MainStreamRepository
    Result.Inject MainStreamPersistence
    Set CreateMainStreamRepository = Result
End Function

Private Function CreateSpecialStreamPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_ISpecialStreamPersistence
    Dim Result As Inf_ISpecialStreamPersistence
    Set Result = New Inf_SpecialStreamPersistence
    Result.Inject CSV
    Set CreateSpecialStreamPersistence = Result
End Function

Private Function CreateSpecialStreamQueryService(ByVal SpecialStreamPersistence As Inf_ISpecialStreamPersistence) As Inf_ISpecialStreamQueryService
    Dim Result As Inf_ISpecialStreamQueryService
    Set Result = New Inf_SpecialStreamQueryService
    Result.Inject SpecialStreamPersistence
    Set CreateSpecialStreamQueryService = Result
End Function

Private Function CreateSpecialStreamRepository(ByVal SpecialStreamPersistence As Inf_ISpecialStreamPersistence) As Inf_ISpecialStreamRepository
    Dim Result As Inf_ISpecialStreamRepository
    Set Result = New Inf_SpecialStreamRepository
    Result.Inject SpecialStreamPersistence
    Set CreateSpecialStreamRepository = Result
End Function

Private Function CreateSubjectPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_ISubjectPersistence
    Dim Result As Inf_ISubjectPersistence
    Set Result = New Inf_ISubjectPersistence
    Result.Inject CSV
    Set CreateSubjectPersistence = Result
End Function

Private Function CreateSubjectQueryService(ByVal SubjectPersistence As Inf_ISubjectPersistence) As Inf_ISubjectQueryService
    Dim Result As Inf_ISubjectQueryService
    Set Result = New Inf_ISubjectQueryService
    Result.Inject SubjectPersistence
    Set CreateSubjectQueryService = Result
End Function

Private Function CreateSubjectRepository(ByVal SubjectPersistence As Inf_ISubjectPersistence) As Inf_ISubjectRepository
    Dim Result As Inf_ISubjectRepository
    Set Result = New Inf_ISubjectRepository
    Result.Inject SubjectPersistence
    Set CreateSubjectRepository = Result
End Function

Private Function CreatePeriodPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_IPeriodPersistence
    Dim Result As Inf_IPeriodPersistence
    Set Result = New Inf_IPeriodPersistence
    Result.Inject CSV
    Set CreatePeriodPersistence = Result
End Function

Private Function CreatePeriodQueryService(ByVal PeriodPersistence As Inf_IPeriodPersistence) As Inf_IPeriodQueryService
    Dim Result As Inf_IPeriodQueryService
    Set Result = New Inf_IPeriodQueryService
    Result.Inject PeriodPersistence
    Set CreatePeriodQueryService = Result
End Function

Private Function CreatePeriodRepository(ByVal PeriodPersistence As Inf_IPeriodPersistence) As Inf_IPeriodRepository
    Dim Result As Inf_IPeriodRepository
    Set Result = New Inf_IPeriodRepository
    Result.Inject PeriodPersistence
    Set CreatePeriodRepository = Result
End Function

Private Function CreateSchoolEventPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_ISchoolEventPersistence
    Dim Result As Inf_ISchoolEventPersistence
    Set Result = New Inf_ISchoolEventPersistence
    Result.Inject CSV
    Set CreateSchoolEventPersistence = Result
End Function

Private Function CreateSchoolEventQueryService(ByVal SchoolEventPersistence As Inf_ISchoolEventPersistence) As Inf_ISchoolEventQueryService
    Dim Result As Inf_ISchoolEventQueryService
    Set Result = New Inf_ISchoolEventQueryService
    Result.Inject SchoolEventPersistence
    Set CreateSchoolEventQueryService = Result
End Function

Private Function CreateSchoolEventRepository(ByVal SchoolEventPersistence As Inf_ISchoolEventPersistence) As Inf_ISchoolEventRepository
    Dim Result As Inf_ISchoolEventRepository
    Set Result = New Inf_ISchoolEventRepository
    Result.Inject SchoolEventPersistence
    Set CreateSchoolEventRepository = Result
End Function
