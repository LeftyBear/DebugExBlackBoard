Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit
Option Private Module

Public Sub Run()
    'ErrorLogger---------------------------------------------------------------------------------
    Dim Logger As App_ILogger
    Set Logger = App_LoggerFactory.CreateLogger
    On Error GoTo ErrorHandler
    'SchoolStructure-----------------------------------------------------------------------------
    Dim GenerateStructure As App_AggregateSchoolStructure
    Set GenerateStructure = App_UseCaseFactory.CreateGenerateSchoolStructure
    Dim Structure As Dom_SchoolStructure
    Set Structure = GenerateStructure.Execute
    'UseCase-------------------------------------------------------------------------------------
    Dim Schedule As App_AggregateSchedule
    Set Schedule = App_UseCaseFactory.CreateAggregateSchedule
    Dim Enrollment As App_AggregateEnrollment
    Set Enrollment = App_UseCaseFactory.CreateAggregateEnrollment
    Dim ClassHour As App_AggregateClassHour
    Set ClassHour = App_UseCaseFactory.CreateAggregateClassHour
    'Presentation--------------------------------------------------------------------------------
    Dim Presenter As Pre_Presenter
    Set Presenter = New Pre_Presenter
    Presenter.Initialize Structure, Logger, Schedule, Enrollment, ClassHour
    Dim Controller As Pre_ClassHourController
    Set Controller = New Pre_ClassHourController
    Controller.Initialize Presenter
    Dim MainView As Pre_MainView
    Set MainView = New Pre_MainView
    MainView.Initialize Controller
    Presenter.AttachView MainView
    MainView.Show vbModeless
    Exit Sub
ErrorHandler:
    Logger.WriteLog Err.Source & vbTab & Err.Description
End Sub
