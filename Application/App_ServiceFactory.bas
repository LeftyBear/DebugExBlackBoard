Attribute VB_Name = "App_ServiceFactory"
'@Folder "Application.CompositionRoot"
Option Explicit

Public Function CreateTotalizationService() As App_TotalizationService
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_TotalizationRepository
    Set BaseRepository = New Inf_TotalizationRepository
    BaseRepository.Initialize ReadRepository
    Dim Service As App_TotalizationService
    Set Service = New App_TotalizationService
    Service.Initialize BaseRepository
    Set CreateTotalizationService = Service
End Function

Public Function CreateLimitValueService() As App_LimitValueService
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_TotalizationRepository
    Set BaseRepository = New Inf_TotalizationRepository
    BaseRepository.Initialize ReadRepository
    Dim Service As App_LimitValueService
    Set Service = New App_LimitValueService
    Service.Initialize BaseRepository
    Set CreateLimitValueService = Service
End Function

Public Function CreateEnrollmentService() As App_EnrollmentService
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_EnrollmentRepository
    Set BaseRepository = New Inf_EnrollmentRepository
    BaseRepository.Initialize ReadRepository
    Dim Service As App_EnrollmentService
    Set Service = New App_EnrollmentService
    Service.Initialize New Dom_SchoolYearCalculator, BaseRepository
    Set CreateEnrollmentService = Service
End Function

Public Function CreateClassHourService() As App_ClassHourService
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_ClassHourRepository
    Set BaseRepository = New Inf_ClassHourRepository
    BaseRepository.Initialize ReadRepository
    Dim Service As App_ClassHourService
    Set Service = New App_ClassHourService
    Service.Initialize New Dom_SchoolYearCalculator, BaseRepository
    Set CreateClassHourService = Service
End Function
