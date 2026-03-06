Attribute VB_Name = "App_UseCaseFactory"
'@Folder "Application.CompositionRoot"
Option Explicit

Public Function CreateTotalizationUseCase() As App_TotalizationUseCase
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_TotalizationRepository
    Set BaseRepository = New Inf_TotalizationRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_TotalizationUseCase
    Set UseCase = New App_TotalizationUseCase
    UseCase.Initialize BaseRepository
    Set CreateTotalizationUseCase = UseCase
End Function

Public Function CreateLimitValueUseCase() As App_LimitValueUseCase
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_TotalizationRepository
    Set BaseRepository = New Inf_TotalizationRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_LimitValueUseCase
    Set UseCase = New App_LimitValueUseCase
    UseCase.Initialize BaseRepository
    Set CreateLimitValueUseCase = UseCase
End Function

Public Function CreateEnrollmentUseCase() As App_EnrollmentUseCase
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_EnrollmentRepository
    Set BaseRepository = New Inf_EnrollmentRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_EnrollmentUseCase
    Set UseCase = New App_EnrollmentUseCase
    UseCase.Initialize New Dom_SchoolYearCalculator, BaseRepository
    Set CreateEnrollmentUseCase = UseCase
End Function

Public Function CreateClassHourUseCase() As App_ClassHourUseCase
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_ClassHourRepository
    Set BaseRepository = New Inf_ClassHourRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_ClassHourUseCase
    Set UseCase = New App_ClassHourUseCase
    UseCase.Initialize New Dom_SchoolYearCalculator, BaseRepository
    Set CreateClassHourUseCase = UseCase
End Function
