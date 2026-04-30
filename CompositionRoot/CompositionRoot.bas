Attribute VB_Name = "CompositionRoot"
'@Folder("CompositionRoot")
Option Explicit
Option Private Module

Public Sub Boot()
    'ErrorLogger ===========================================================
    Dim FilePath As String
    FilePath = BuildLogFilePath
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "CompositionRoot", "File path is invalid."
    Dim Logger As App_ILogger
    Set Logger = CreateLogger(FilePath)
    On Error GoTo ErrHandle
    'UseCase ===============================================================
    Dim Structure As App_SchoolStructureAggregator
    Set Structure = BuildStructureUseCase
    Dim ClassHour As App_ClassHourAggregator
    Set ClassHour = BuildClassHourUseCase
    Dim Enrollment As App_EnrollmentAggregator
    Set Enrollment = BuildEnrollmentUseCase
    Dim Schedule As App_ScheduleAggregator
    Set Schedule = BuildScheduleUseCase
    Dim ViewBuilder As App_ViewDTOBuilder
    Set ViewBuilder = BuildViewDTOUseCase
    'Presentation ==========================================================
    Dim Presenter As Pre_CalenderPresenter
    Set Presenter = New Pre_CalenderPresenter
    Presenter.Initialize Structure, Logger, ClassHour, Enrollment, Schedule, _
        New Dom_ValueObjectFactory, New Pre_BasePresenter, ViewBuilder
    Dim Controller As Pre_CalenderController
    Set Controller = New Pre_CalenderController
    Controller.Initialize Presenter
    Dim MainView As Pre_MainView
    Set MainView = New Pre_MainView
    MainView.Initialize Controller
    Presenter.AttachView MainView
    MainView.Show vbModeless
    Exit Sub
ErrHandle:
    Logger.WriteLog Err.Source & vbTab & Err.Description
End Sub

'ErrorLogger ===============================================================
Private Function BuildLogFilePath() As String
    Dim Builder As App_LogFilePathBuilder
    Set Builder = New App_LogFilePathBuilder
    BuildLogFilePath = Builder.Builed
End Function

Private Function CreateLogger(ByVal FilePath As String) As App_ILogger
    Dim TypeCode As Inf_EnvironmentTypePolicy
    TypeCode = Inf_TypePolicy.GetEnvironmentTypeCode
    Dim Selector As App_LoggerSelector
    Set Selector = New App_LoggerSelector
    Set CreateLogger = Selector.SelectLogger(TypeCode, New Inf_FileLogger, New Inf_DebugLogger)
End Function

'UseCase ===================================================================
Private Function BuildStructureUseCase() As App_SchoolStructureAggregator
    'Reader ----------------------------------------------------------------
    Dim Reader As App_ICSVReader
    Set Reader = New Inf_CSVReader
    'Builder ---------------------------------------------------------------
    Dim Builder As App_ConfigFilePathBuilder
    Set Builder = New App_ConfigFilePathBuilder
    'UseCase ---------------------------------------------------------------
    Dim UseCase As App_SchoolStructureAggregator
    Set UseCase = New App_SchoolStructureAggregator
    UseCase.Initialize Reader, Builder, New App_UpperValueFactory, _
        New App_SpecialStreamCatalogFactory, New App_SubjectCatalogFactory, _
        New App_SchoolEventCatalogFactory
    Set BuildStructureUseCase = UseCase
End Function

Private Function BuildClassHourUseCase() As App_ClassHourAggregator
    'Reader ----------------------------------------------------------------
    Dim Reader As App_ICSVReader
    Set Reader = New Inf_CSVReader
    'Builder ---------------------------------------------------------------
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    'Resolver --------------------------------------------------------------
    Dim ColResolver As App_ClassHourColumnResolver
    Set ColResolver = New App_ClassHourColumnResolver
    Dim RowResolver As App_ClassHourRowResolver
    Set RowResolver = New App_ClassHourRowResolver
    'Factory ---------------------------------------------------------------
    Dim Factory As App_ClassHourYearAggFactory
    Set Factory = CreateClassHourYearAggFactory
    'UseCase ---------------------------------------------------------------
    Dim UseCase As App_ClassHourAggregator
    Set UseCase = New App_ClassHourAggregator
    UseCase.Initialize Reader, Builder, ColResolver, RowResolver, Factory
    Set BuildClassHourUseCase = UseCase
End Function

Private Function BuildEnrollmentUseCase() As App_EnrollmentAggregator
    'Reader ----------------------------------------------------------------
    Dim Reader As App_ICSVReader
    Set Reader = New Inf_CSVReader
    'Builder ---------------------------------------------------------------
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    'Resolver --------------------------------------------------------------
    Dim ColResolver As App_EnrollmentColumnResolver
    Set ColResolver = New App_EnrollmentColumnResolver
    Dim RowResolver As App_EnrollmentRowResolver
    Set RowResolver = New App_EnrollmentRowResolver
    'Factory ---------------------------------------------------------------
    Dim Factory As App_EnrollmentYearAggFactory
    Set Factory = CreateEnrollmentYearAggFactory
    'UseCase ---------------------------------------------------------------
    Dim UseCase As App_EnrollmentAggregator
    Set UseCase = New App_EnrollmentAggregator
    UseCase.Initialize Reader, Builder, ColResolver, RowResolver, Factory
    Set BuildEnrollmentUseCase = UseCase
End Function

Private Function BuildScheduleUseCase() As App_ScheduleAggregator
    'Reader ----------------------------------------------------------------
    Dim Reader As App_ICSVReader
    Set Reader = New Inf_CSVReader
    'Builder ---------------------------------------------------------------
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    'Resolver --------------------------------------------------------------
    Dim ColResolver As App_ScheduleColumnResolver
    Set ColResolver = New App_ScheduleColumnResolver
    Dim RowResolver As App_ScheduleRowResolver
    Set RowResolver = New App_ScheduleRowResolver
    'Factory ---------------------------------------------------------------
    Dim Factory As App_ScheduleYearAggFactory
    Set Factory = CreateScheduleYearAggFactory
    'UseCase ---------------------------------------------------------------
    Dim UseCase As App_ScheduleAggregator
    Set UseCase = New App_ScheduleAggregator
    UseCase.Initialize Reader, Builder, ColResolver, RowResolver, Factory
    Set BuildScheduleUseCase = UseCase
End Function

Private Function BuildViewDTOUseCase() As App_ViewDTOBuilder
    Dim ClassHourFormatter As App_ClassHourFormatter
    Set ClassHourFormatter = CreateClassHourFormatter
    Dim EnrollmentFormatter As App_EnrollmentFormatter
    Set EnrollmentFormatter = CreateEnrollmentFormatter
    Dim TimeTableFormatter As App_TimeTableFormatter
    Set TimeTableFormatter = CreateTimeTableFormatter
    Dim Builder As App_ViewDTOBuilder
    Set Builder = New App_ViewDTOBuilder
    Builder.Initialize ClassHourFormatter, EnrollmentFormatter, TimeTableFormatter
    Set BuildViewDTOUseCase = Builder
End Function

'Factory -------------------------------------------------------------------
Private Function CreateClassHourYearAggFactory() As App_ClassHourYearAggFactory
    Dim ClassHourFactory As Dom_ClassHourFactory
    Set ClassHourFactory = CreateClassHourFactory
    Dim TimeTableFactory As Dom_TimeTableFactory
    Set TimeTableFactory = CreateTimeTableFactory
    Dim Result As App_ClassHourYearAggFactory
    Set Result = New App_ClassHourYearAggFactory
    Result.Initialize ClassHourFactory, TimeTableFactory
    Set CreateClassHourYearAggFactory = Result
End Function

Private Function CreateClassHourFactory() As Dom_ClassHourFactory
    Dim Result As Dom_ClassHourFactory
    Set Result = New Dom_ClassHourFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateClassHourFactory = Result
End Function

Private Function CreateTimeTableFactory() As Dom_TimeTableFactory
    Dim Result As Dom_TimeTableFactory
    Set Result = New Dom_TimeTableFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateTimeTableFactory = Result
End Function

Private Function CreateEnrollmentYearAggFactory() As App_EnrollmentYearAggFactory
    Dim Factory As Dom_EnrollmentFactory
    Set Factory = CreateEnrollmentFactory
    Dim Result As App_EnrollmentYearAggFactory
    Set Result = New App_EnrollmentYearAggFactory
    Result.Initialize Factory
    Set CreateEnrollmentYearAggFactory = Result
End Function

Private Function CreateEnrollmentFactory() As Dom_EnrollmentFactory
    Dim Result As Dom_EnrollmentFactory
    Set Result = New Dom_EnrollmentFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateEnrollmentFactory = Result
End Function

Private Function CreateScheduleYearAggFactory() As App_ScheduleYearAggFactory
    Dim Factory As Dom_ScheduleFactory
    Set Factory = CreateScheduleFactory
    Dim Result As App_ScheduleYearAggFactory
    Set Result = New App_ScheduleYearAggFactory
    Result.Initialize Factory
    Set CreateScheduleYearAggFactory = Result
End Function

Private Function CreateScheduleFactory() As Dom_ScheduleFactory
    Dim Result As Dom_ScheduleFactory
    Set Result = New Dom_ScheduleFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateScheduleFactory = Result
End Function

'Service -------------------------------------------------------------------
Private Function CreateClassHourFormatter() As App_ClassHourFormatter
    Dim Service As App_ClassHourSummaryCalculator
    Set Service = CreateClassHourService
    Dim Result As App_ClassHourFormatter
    Set Result = New App_ClassHourFormatter
    Result.Initialize Service, New Dom_ValueObjectFactory
    Set CreateClassHourFormatter = Result
End Function

Private Function CreateEnrollmentFormatter() As App_EnrollmentFormatter
    Dim Service As App_EnrollmentSummaryCalculator
    Set Service = CreateEnrollmentService
    Dim Result As App_EnrollmentFormatter
    Set Result = New App_EnrollmentFormatter
    Result.Initialize Service, New Dom_ValueObjectFactory
    Set CreateEnrollmentFormatter = Result
End Function

Private Function CreateTimeTableFormatter() As App_TimeTableFormatter
    Dim Service As App_TimeTableService
    Set Service = CreateTimeTableService
    Dim Result As App_TimeTableFormatter
    Set Result = New App_TimeTableFormatter
    Result.Initialize Service, New Dom_ValueObjectFactory
    Set CreateTimeTableFormatter = Result
End Function

Private Function CreateClassHourService() As App_ClassHourSummaryCalculator
    Dim Calculator As App_ClassHourCalculator
    Set Calculator = CreateClassHourCalculator
    Dim Result As App_ClassHourSummaryCalculator
    Set Result = New App_ClassHourSummaryCalculator
    Result.Initialize Calculator
    Set CreateClassHourService = Result
End Function

Private Function CreateClassHourCalculator() As App_ClassHourCalculator
    Dim Result As App_ClassHourCalculator
    Set Result = New App_ClassHourCalculator
    Result.Initialize New App_ClassHourFilter
    Set CreateClassHourCalculator = Result
End Function

Private Function CreateEnrollmentService() As App_EnrollmentSummaryCalculator
    Dim Calculator As App_EnrollmentCalculator
    Set Calculator = CreateEnrollmentCalculator
    Dim Result As App_EnrollmentSummaryCalculator
    Set Result = New App_EnrollmentSummaryCalculator
    Result.Initialize Calculator
    Set CreateEnrollmentService = Result
End Function

Private Function CreateEnrollmentCalculator() As App_EnrollmentCalculator
    Dim Result As App_EnrollmentCalculator
    Set Result = New App_EnrollmentCalculator
    Result.Initialize New App_EnrollmentFilter
    Set CreateEnrollmentCalculator = Result
End Function

Private Function CreateTimeTableService() As App_TimeTableService
    Dim Counter As App_TimeTableCounter
    Set Counter = CreateTimeTableCounter
    Dim Result As App_TimeTableService
    Set Result = New App_TimeTableService
    Result.Initialize Counter
    Set CreateTimeTableService = Result
End Function

Private Function CreateTimeTableCounter() As App_TimeTableCounter
    Dim Result As App_TimeTableCounter
    Set Result = New App_TimeTableCounter
    Result.Initialize New App_TimeTableFilter
    Set CreateTimeTableCounter = Result
End Function
