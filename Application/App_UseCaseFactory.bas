Attribute VB_Name = "App_UseCaseFactory"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Function CreateGenerateSchoolStructure() As App_AggregateSchoolStructure
    'Builder------------------------------------------------------------------------------------
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_ConfigFilePathBuilder
    Set Builder = New App_ConfigFilePathBuilder
    Builder.Initialize Provider
    'Repository---------------------------------------------------------------------------------
    Dim Repository As Inf_ConfigRepository
    Set Repository = New Inf_ConfigRepository
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateSchoolStructure
    Set UseCase = New App_AggregateSchoolStructure
    UseCase.Initialize Builder, Repository
    Set CreateGenerateSchoolStructure = UseCase
End Function

Public Function CreateAggregateEnrollment(ByVal Structure As Dom_SchoolStructure) As App_AggregateEnrollment
    'Builder------------------------------------------------------------------------------------
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    Builder.Initialize Provider
    'Repository---------------------------------------------------------------------------------
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    Dim BaseRepository As Inf_EnrollmentRepository
    Set BaseRepository = New Inf_EnrollmentRepository
    BaseRepository.Initialize ReadRepository
    'Resolver-----------------------------------------------------------------------------------
    Dim Resolver As Dom_EnrollmentColumnResolver
    Set Resolver = New Dom_EnrollmentColumnResolver
    Resolver.Initialize Structure
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateEnrollment
    Set UseCase = New App_AggregateEnrollment
    UseCase.Initialize Builder, BaseRepository, Resolver
    Set CreateAggregateEnrollment = UseCase
End Function

Public Function CreateAggregateClassHour(ByVal Structure As Dom_SchoolStructure) As App_AggregateClassHour
    'Builder------------------------------------------------------------------------------------
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    Builder.Initialize Provider
    'Repository---------------------------------------------------------------------------------
    Dim ReadRepository As Inf_CSVReadRepository
    Set ReadRepository = New Inf_CSVReadRepository
    Dim BaseRepository As Inf_ClassHourRepository
    Set BaseRepository = New Inf_ClassHourRepository
    BaseRepository.Initialize ReadRepository
    'Resolver-----------------------------------------------------------------------------------
    Dim Resolver As Dom_ClassHourColumnResolver
    Set Resolver = New Dom_ClassHourColumnResolver
    Resolver.Initialize Structure
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateClassHour
    Set UseCase = New App_AggregateClassHour
    UseCase.Initialize Builder, BaseRepository, Resolver
    Set CreateAggregateClassHour = UseCase
End Function
