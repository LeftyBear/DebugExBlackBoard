Attribute VB_Name = "App_UseCaseFactory"
'@Folder "Application.Factory"
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
    Dim Repository As Inf_ConfigCSVReader
    Set Repository = New Inf_ConfigCSVReader
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateSchoolStructure
    Set UseCase = New App_AggregateSchoolStructure
    UseCase.Initialize Builder, Repository
    Set CreateGenerateSchoolStructure = UseCase
End Function

Public Function CreateAggregateSchedule() As App_AggregateSchedule
    'Builder------------------------------------------------------------------------------------
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    Builder.Initialize Provider
    'Repository---------------------------------------------------------------------------------
    Dim Repository As Inf_CSVReadRepository
    Set Repository = New Inf_CSVReadRepository
    Dim Reader As Inf_ScheduleCSVReader
    Set Reader = New Inf_ScheduleCSVReader
    Reader.Initialize Repository
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateSchedule
    Set UseCase = New App_AggregateSchedule
    UseCase.Initialize Builder, Reader
    Set CreateAggregateSchedule = UseCase
End Function

Public Function CreateAggregateEnrollment() As App_AggregateEnrollment
    'Builder------------------------------------------------------------------------------------
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    Builder.Initialize Provider
    'Repository---------------------------------------------------------------------------------
    Dim Repository As Inf_CSVReadRepository
    Set Repository = New Inf_CSVReadRepository
    Dim Reader As Inf_EnrollmentCSVReader
    Set Reader = New Inf_EnrollmentCSVReader
    Reader.Initialize Repository
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateEnrollment
    Set UseCase = New App_AggregateEnrollment
    UseCase.Initialize Builder, Reader
    Set CreateAggregateEnrollment = UseCase
End Function

Public Function CreateAggregateClassHour() As App_AggregateClassHour
    'Builder------------------------------------------------------------------------------------
    Dim Provider As App_IWorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Builder As App_EntityFilePathBuilder
    Set Builder = New App_EntityFilePathBuilder
    Builder.Initialize Provider
    'Repository---------------------------------------------------------------------------------
    Dim Repository As Inf_CSVReadRepository
    Set Repository = New Inf_CSVReadRepository
    Dim Reader As Inf_ClassHourCSVReader
    Set Reader = New Inf_ClassHourCSVReader
    Reader.Initialize Repository
    'UseCase------------------------------------------------------------------------------------
    Dim UseCase As App_AggregateClassHour
    Set UseCase = New App_AggregateClassHour
    UseCase.Initialize Builder, Reader
    Set CreateAggregateClassHour = UseCase
End Function
