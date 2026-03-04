Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit

Public Sub Run()
    Dim TotalizationService As App_TotalizationService
    Set TotalizationService = App_ServiceFactory.CreateTotalizationService
    TotalizationService.Excute
    Dim LimitValueService As App_LimitValueService
    Set LimitValueService = App_ServiceFactory.CreateLimitValueService
    LimitValueService.Excute
    
    Dim EnrollmentService As App_EnrollmentService
    Set EnrollmentService = App_ServiceFactory.CreateEnrollmentService
    EnrollmentService.Excute Date
End Sub
