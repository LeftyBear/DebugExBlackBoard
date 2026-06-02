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
    Dim ClassHourPersistence As Inf_ClassHourPersistence
    Set ClassHourPersistence =CreateClassHourPersistence(CSV)
    Dim EnrollmentPersistence As Inf_EnrollmentPersistence
    Set EnrollmentPersistence = CreateEnrollmentPersistence(CSV)
    Dim SchedulePersistence As Inf_SchedulePersistence
    Set SchedulePersistence = CreateSchedulePersistence(CSV)
    Dim MainStreamPersistence As Inf_MainStreamPersistence
    Set MainStreamPersistence = CreateMainStreamPersistence(CSV)
    Dim SpecialStreamPersistence As Inf_SpecialStreamPersistence
    Set SpecialStreamPersistence = CreateSpecialStreamPersistence(CSV)
    Dim SubjectPersistence As Inf_SubjectPersistence
    Set SubjectPersistence = CreateSubjectPersistence(CSV)
    Dim PeriodPersistence As Inf_PeriodPersistence
    Set PeriodPersistence = CreatePeriodPersistence(CSV)
    Dim SchoolEventPersistence As Inf_SchoolEventPersistence
    Set SchoolEventPersistence = CreateSchoolEventPersistence(CSV)
    'QueryService ----------------------------------------------------------
    Dim ClassHourQueryService As Inf_ClassHourQueryService
    Set ClassHourQueryService =CreateClassHourQueryService(ClassHourPersistence)
    Dim EnrollmentQueryService As Inf_EnrollmentQueryService
    Set EnrollmentQueryService = CreateEnrollmentQueryService(EnrollmentPersistence)
    Dim ScheduleQueryService As Inf_ScheduleQueryService
    Set ScheduleQueryService = CreateScheduleQueryService(SchedulePersistence)
    Dim MainStreamQueryService As Inf_MainStreamQueryService
    Set MainStreamQueryService = CreateMainStreamQueryService(MainStreamPersistence)
    Dim SpecialStreamQueryService As Inf_SpecialStreamQueryService
    Set SpecialStreamQueryService = CreateSpecialStreamQueryService(SpecialStreamPersistence)
    Dim SubjectQueryService As Inf_SubjectQueryService
    Set SubjectQueryService = CreateSubjectQueryService(SubjectPersistence)
    Dim PeriodQueryService As Inf_PeriodQueryService
    Set PeriodQueryService = CreatePeriodQueryService(PeriodPersistence)
    Dim SchoolEventQueryService As Inf_SchoolEventQueryService
    Set SchoolEventQueryService = CreateSchoolEventQueryService(SchoolEventPersistence)
    'Repository ------------------------------------------------------------
    Dim ClassHourRepository As Inf_ClassHourRepository
    Set ClassHourRepository = CreateClassHourRepository(ClassHourPersistence)
    Dim EnrollmentRepository As Inf_EnrollmentRepository
    Set EnrollmentRepository = CreateEnrollmentRepository(EnrollmentPersistence)
    Dim ScheduleRepository As Inf_ScheduleRepository
    Set ScheduleRepository = CreateScheduleRepository(SchedulePersistence)
    Dim MainStreamRepository As Inf_MainStreamRepository
    Set MainStreamRepository = CreateMainStreamRepository(MainStreamPersistence)
    Dim SpecialStreamRepository As Inf_SpecialStreamRepository
    Set SpecialStreamRepository = CreateSpecialStreamRepository(SpecialStreamPersistence)
    Dim SubjectRepository As Inf_SubjectRepository
    Set SubjectRepository = CreateSubjectRepository(SubjectPersistence)
    Dim PeriodRepository As Inf_PeriodRepository
    Set PeriodRepository = CreatePeriodRepository(PeriodPersistence)
    Dim SchoolEventRepository As Inf_SchoolEventRepository
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

Private Function CreateClassHourPersistence(Byval CSV as Inf_CSVPersistence) As Inf_ClassHourParsistence
    Dim Result As Inf_ClassHourParsistence
    Set Result = New Inf_ClassHourParsistence
    Result.Inject CSV
    Set CreateClassHourPersistence = Result
End Function

Private Function CreateClassHourQueryService(ByVal ClassHourPersistence As Inf_ClassHourPersistence) As Inf_ClassHourQueryService
    Dim Result As Inf_ClassHourQueryService
    Set Result = New Inf_ClassHourQueryService
    Result.Inject ClassHourPersistence
    Set CreateClassHourQueryService = Result
End Function

Private Function CreateClassHourRepository(ByVal ClassHourPersistence As Inf_ClassHourPersistence) As Inf_ClassHourRepository
Dim Result As Inf_ClassHourRepository
    Set Result = New Inf_ClassHourRepository
    Result.Inject ClassHourPersistence
    Set CreateClassHourRepository = Result
End Function

Private Function CreateEnrollmentPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_EnrollmentPersistence
    Dim Result As Inf_EnrollmentPersistence
    Set Result = New Inf_EnrollmentPersistence
    Result.Inject CSV
    Set CreateEnrollmentPersistence = Result
End Function

Private Function CreateEnrollmentQueryService(ByVal EnrollmentPersistence As Inf_EnrollmentPersistence) As Inf_EnrollmentQueryService
    Dim Result As Inf_EnrollmentQueryService
    Set Result = New Inf_EnrollmentQueryService
    Result.Inject EnrollmentPersistence
    Set CreateEnrollmentQueryService = Result
End Function

Private Function CreateEnrollmentRepository(ByVal EnrollmentPersistence As Inf_EnrollmentPersistence) As Inf_EnrollmentRepository
    Dim Result As Inf_EnrollmentRepository
    Set Result = New Inf_EnrollmentRepository
    Result.Inject EnrollmentPersistence
    Set CreateEnrollmentRepository = Result
End Function

Private Function CreateSchedulePersistence(ByVal CSV As Inf_CSVPersistence) As Inf_SchedulePersistence
    Dim Result As Inf_SchedulePersistence
    Set Result = New Inf_SchedulePersistence
    Result.Inject CSV
    Set CreateSchedulePersistence = Result
End Function

Private Function CreateScheduleQueryService(ByVal SchedulePersistence As Inf_SchedulePersistence) As Inf_ScheduleQueryService
    Dim Result As Inf_ScheduleQueryService
    Set Result = New Inf_ScheduleQueryService
    Result.Inject SchedulePersistence
    Set CreateScheduleQueryService = Result
End Function

Private Function CreateScheduleRepository(ByVal SchedulePersistence As Inf_SchedulePersistence) As Inf_ScheduleRepository
    Dim Result As Inf_ScheduleRepository
    Set Result = New Inf_ScheduleRepository
    Result.Inject SchedulePersistence
    Set CreateScheduleRepository = Result
End Function

Private Function CreateMainStreamPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_MainStreamPersistence
    Dim Result As Inf_MainStreamPersistence
    Set Result = New Inf_MainStreamPersistence
    Result.Inject CSV
    Set CreateMainStreamPersistence = Result
End Function

Private Function CreateMainStreamQueryService(ByVal MainStreamPersistence As Inf_MainStreamPersistence) As Inf_MainStreamQueryService
    Dim Result As Inf_MainStreamQueryService
    Set Result = New Inf_MainStreamQueryService
    Result.Inject MainStreamPersistence
    Set CreateMainStreamQueryService = Result
End Function

Private Function CreateMainStreamRepository(ByVal MainStreamPersistence As Inf_MainStreamPersistence) As Inf_MainStreamRepository
    Dim Result As Inf_MainStreamRepository
    Set Result = New Inf_MainStreamRepository
    Result.Inject MainStreamPersistence
    Set CreateMainStreamRepository = Result
End Function

Private Function CreateSpecialStreamPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_SpecialStreamPersistence
    Dim Result As Inf_SpecialStreamPersistence
    Set Result = New Inf_SpecialStreamPersistence
    Result.Inject CSV
    Set CreateSpecialStreamPersistence = Result
End Function

Private Function CreateSpecialStreamQueryService(ByVal SpecialStreamPersistence As Inf_SpecialStreamPersistence) As Inf_SpecialStreamQueryService
    Dim Result As Inf_SpecialStreamQueryService
    Set Result = New Inf_SpecialStreamQueryService
    Result.Inject SpecialStreamPersistence
    Set CreateSpecialStreamQueryService = Result
End Function

Private Function CreateSpecialStreamRepository(ByVal SpecialStreamPersistence As Inf_SpecialStreamPersistence) As Inf_SpecialStreamRepository
    Dim Result As Inf_SpecialStreamRepository
    Set Result = New Inf_SpecialStreamRepository
    Result.Inject SpecialStreamPersistence
    Set CreateSpecialStreamRepository = Result
End Function

Private Function CreateSubjectPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_SubjectPersistence
    Dim Result As Inf_SubjectPersistence
    Set Result = New Inf_SubjectPersistence
    Result.Inject CSV
    Set CreateSubjectPersistence = Result
End Function

Private Function CreateSubjectQueryService(ByVal SubjectPersistence As Inf_SubjectPersistence) As Inf_SubjectQueryService
    Dim Result As Inf_SubjectQueryService
    Set Result = New Inf_SubjectQueryService
    Result.Inject SubjectPersistence
    Set CreateSubjectQueryService = Result
End Function

Private Function CreateSubjectRepository(ByVal SubjectPersistence As Inf_SubjectPersistence) As Inf_SubjectRepository
    Dim Result As Inf_SubjectRepository
    Set Result = New Inf_SubjectRepository
    Result.Inject SubjectPersistence
    Set CreateSubjectRepository = Result
End Function

Private Function CreatePeriodPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_PeriodPersistence
    Dim Result As Inf_PeriodPersistence
    Set Result = New Inf_PeriodPersistence
    Result.Inject CSV
    Set CreatePeriodPersistence = Result
End Function

Private Function CreatePeriodQueryService(ByVal PeriodPersistence As Inf_PeriodPersistence) As Inf_PeriodQueryService
    Dim Result As Inf_PeriodQueryService
    Set Result = New Inf_PeriodQueryService
    Result.Inject PeriodPersistence
    Set CreatePeriodQueryService = Result
End Function

Private Function CreatePeriodRepository(ByVal PeriodPersistence As Inf_PeriodPersistence) As Inf_PeriodRepository
    Dim Result As Inf_PeriodRepository
    Set Result = New Inf_PeriodRepository
    Result.Inject PeriodPersistence
    Set CreatePeriodRepository = Result
End Function

Private Function CreateSchoolEventPersistence(ByVal CSV As Inf_CSVPersistence) As Inf_SchoolEventPersistence
    Dim Result As Inf_SchoolEventPersistence
    Set Result = New Inf_SchoolEventPersistence
    Result.Inject CSV
    Set CreateSchoolEventPersistence = Result
End Function

Private Function CreateSchoolEventQueryService(ByVal SchoolEventPersistence As Inf_SchoolEventPersistence) As Inf_SchoolEventQueryService
    Dim Result As Inf_SchoolEventQueryService
    Set Result = New Inf_SchoolEventQueryService
    Result.Inject SchoolEventPersistence
    Set CreateSchoolEventQueryService = Result
End Function

Private Function CreateSchoolEventRepository(ByVal SchoolEventPersistence As Inf_SchoolEventPersistence) As Inf_SchoolEventRepository
    Dim Result As Inf_SchoolEventRepository
    Set Result = New Inf_SchoolEventRepository
    Result.Inject SchoolEventPersistence
    Set CreateSchoolEventRepository = Result
End Function
