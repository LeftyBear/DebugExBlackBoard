Attribute VB_Name = "App_UseCaseFactory"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Function CreateSchoolConfigGenerater() As App_GenerateSchoolStructure
    Dim PathBuilder As Inf_ConfigFilePathBuilder
    Set PathBuilder = New Inf_ConfigFilePathBuilder
    PathBuilder.Initialize New Inf_ConfigFileNameResolver, New Inf_WorkbookPathProvider
    Dim ReadRepository As Inf_ConfigReadRepository
    Set ReadRepository = New Inf_ConfigReadRepository
    ReadRepository.Initialize PathBuilder, New Inf_TextStreamReader, New Inf_CSVRFCParser
    Dim BaseRepository As Inf_ConfigRepository
    Set BaseRepository = New Inf_ConfigRepository
    BaseRepository.Initialize ReadRepository
    Dim UseCase As App_GenerateSchoolStructure
    Set UseCase = New App_GenerateSchoolStructure
    UseCase.Initialize BaseRepository
    Set CreateSchoolConfigGenerater = UseCase
End Function

Public Function CreateAggregateEnrollment(ByVal SchoolStructure As Dom_SchoolStructure) As App_AggregateEnrollment
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
    Dim UseCase As App_AggregateEnrollment
    Set UseCase = New App_AggregateEnrollment
    UseCase.Initialize BaseRepository, SchoolStructure, New Dom_EnrollmentHeaderParser, Interpreter
    Set CreateAggregateEnrollment = UseCase
End Function

Public Function CreateAggregateClassHour(ByVal SchoolStructure As Dom_SchoolStructure) As App_AggregateClassHour
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
    Dim UseCase As App_AggregateClassHour
    Set UseCase = New App_AggregateClassHour
    UseCase.Initialize BaseRepository, SchoolStructure, New Dom_ClassHourHeaderParser, Interpreter
    Set CreateAggregateClassHour = UseCase
End Function
