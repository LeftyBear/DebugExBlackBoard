VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_CompositionRoot_Enrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.CompositionRoot

Option Explicit

Public Function BuildUseCase() As App_EnrollmentUseCase
    Dim StructureResolver As App_SchoolStructureResolver
    Set StructureResolver = New App_SchoolStructureResolver
    Dim YearFolderResolver As App_YearFolderResolver
    Set YearFolderResolver = New App_YearFolderResolver
    Dim FileNameResolver As App_FileNameResolver
    Set FileNameResolver = New App_FileNameResolver
    Dim Repository As Inf_EnrollmentRepository
    Set Repository = New Inf_EnrollmentRepository
    Dim SummaryService As Dom_EnrollmentSummaryService
    Set SummaryService = New Dom_EnrollmentSummaryService
    Dim UseCase As App_EnrollmentUseCase
    Set UseCase = New App_EnrollmentUseCase
    UseCase.Initialize StructureResolver, YearFolderResolver, FileNameResolver, Repository, SummaryService
    Set BuildUseCase = UseCase
End Function
