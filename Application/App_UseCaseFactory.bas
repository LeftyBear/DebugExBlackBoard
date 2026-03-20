Attribute VB_Name = "App_UseCaseFactory"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Function CreateSchoolConfigGenerater() As App_GenerateSchoolConfigUseCase
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_ConfigRepository
    Set BaseRepository = New Inf_ConfigRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_GenerateSchoolConfigUseCase
    Set UseCase = New App_GenerateSchoolConfigUseCase
    UseCase.Initialize BaseRepository
    Set CreateSchoolConfigGenerater = UseCase
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
    Dim Interpreter As Dom_EnrollmentInterpreter
    Set Interpreter = New Dom_EnrollmentInterpreter
    Interpreter.Initialize New Dom_HeaderTokenResolver
    Dim UseCase As App_AggregateEnrollmentUseCase
    Set UseCase = New App_AggregateEnrollmentUseCase
    UseCase.Initialize BaseRepository, SchoolConfig, New Dom_EnrollmentHeaderParser, Interpreter
    Set CreateAggregateEnrollmentUseCase = UseCase
End Function

Public Function CreateAggregateClassHourUseCase(ByVal SchoolConfig As Dom_SchoolConfig) As App_AggregateClassHourUseCase
    Dim PathBuilder As Inf_EntityFilePathBuilder
    Set PathBuilder = New Inf_EntityFilePathBuilder
    PathBuilder.Initialize New Inf_EntityFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_ClassHourRepository
    Set BaseRepository = New Inf_ClassHourRepository
    BaseRepository.Initialize ReadRepository
    Dim Interpreter As Dom_ClassHourInterpreter
    Set Interpreter = New Dom_ClassHourInterpreter
    Interpreter.Initialize New Dom_HeaderTokenResolver
    Dim UseCase As App_AggregateClassHourUseCase
    Set UseCase = New App_AggregateClassHourUseCase
    UseCase.Initialize BaseRepository, SchoolConfig, New Dom_ClassHourHeaderParser, Interpreter
    Set CreateAggregateClassHourUseCase = UseCase
End Function
