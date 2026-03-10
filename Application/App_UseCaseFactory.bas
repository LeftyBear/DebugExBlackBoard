Attribute VB_Name = "App_UseCaseFactory"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Function CreateAggregateSubjectUseCase() As App_AggregateSubjectUseCase
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_SubjectRepository
    Set BaseRepository = New Inf_SubjectRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_AggregateSubjectUseCase
    Set UseCase = New App_AggregateSubjectUseCase
    UseCase.Initialize BaseRepository
    Set CreateAggregateSubjectUseCase = UseCase
End Function

Public Function CreateAggregateLimitValueUseCase() As App_AggregateLimitValueUseCase
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_LimitValueRepository
    Set BaseRepository = New Inf_LimitValueRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_AggregateLimitValueUseCase
    Set UseCase = New App_AggregateLimitValueUseCase
    UseCase.Initialize BaseRepository
    Set CreateAggregateLimitValueUseCase = UseCase
End Function

Public Function CreateAggregateEnrollmentUseCase(ByVal SchoolConfig As Dom_SchoolConfig) As App_AggregateEnrollmentUseCase
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_EnrollmentRepository
    Set BaseRepository = New Inf_EnrollmentRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_AggregateEnrollmentUseCase
    Set UseCase = New App_AggregateEnrollmentUseCase
    UseCase.Initialize New Dom_SchoolYearCalculator, BaseRepository, SchoolConfig
    Set CreateAggregateEnrollmentUseCase = UseCase
End Function

Public Function CreateAggregateClassHourUseCase() As App_AggregateClassHourUseCase
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_ClassHourRepository
    Set BaseRepository = New Inf_ClassHourRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_AggregateClassHourUseCase
    Set UseCase = New App_AggregateClassHourUseCase
    UseCase.Initialize New Dom_SchoolYearCalculator, BaseRepository
    Set CreateAggregateClassHourUseCase = UseCase
End Function
