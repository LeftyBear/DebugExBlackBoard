Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Sub Run()
    'ErrorLog------------------------------------------------------------------------------------
    Dim Logger As Inf_ILogger
    Set Logger = App_LoggerFactory.CreateLogger
    On Error GoTo ErrorHandler
    'SchoolStructure-----------------------------------------------------------------------------
    Dim GenerateStructure As App_AggregateSchoolStructure
    Set GenerateStructure = App_UseCaseFactory.CreateGenerateSchoolStructure
    Dim Structure As Dom_SchoolStructure
    Set Structure = GenerateStructure.Execute
    'UseCase-------------------------------------------------------------------------------------
    Dim AggregateEnrollment As App_AggregateEnrollment
    Set AggregateEnrollment = App_UseCaseFactory.CreateAggregateEnrollment
    Dim AggregateClassHour As App_AggregateClassHour
    Set AggregateClassHour = App_UseCaseFactory.CreateAggregateClassHour
    'Presenter-----------------------------------------------------------------------------------
    Dim Presenter As App_Presenter
    Set Presenter = New App_Presenter
    Presenter.Initialize Logger, Structure, AggregateEnrollment, AggregateClassHour
    'View----------------------------------------------------------------------------------------
    Dim MainView As App_MainView
    Set MainView = New App_MainView
    MainView.Initialize Presenter
    MainView.Show vbModeless
    Exit Sub
ErrorHandler:
    Logger.WriteLog Err.Source & vbTab & Err.Description
End Sub
