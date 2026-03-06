Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
Option Explicit

Public Sub Run()
    Dim TotalizationUseCase As App_TotalizationUseCase
    Set TotalizationUseCase = App_UseCaseFactory.CreateTotalizationUseCase
    TotalizationUseCase.Excute
    Dim LimitValueUseCase As App_LimitValueUseCase
    Set LimitValueUseCase = App_UseCaseFactory.CreateLimitValueUseCase
    LimitValueUseCase.Excute
    
    Dim EnrollmentUseCase As App_EnrollmentUseCase
    Set EnrollmentUseCase = App_UseCaseFactory.CreateEnrollmentUseCase
    EnrollmentUseCase.Excute Date
End Sub
