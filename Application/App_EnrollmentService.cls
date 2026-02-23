VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Load(ByVal FilePath As String) As Dom_EnrollmentAggregate
    If VBA.Len(FilePath) = 0 Then Err.Raise errEmptyFilePath, , "FilePath is empty."
    Dim Repository As Dom_IEnrollmentRepository
    Set Repository = CreateEnrollmentRepository
    Dim Lines() As String
    Lines = Repository.LoadLines(FilePath)
    Dim Loader As App_CSVLoader
    Set Loader = New App_CSVLoader
    Dim RowsOfColumnDtos As VBA.Collection
    Set RowsOfColumnDtos = Loader.Parse(Lines)
    Dim Builder As App_EnrollmentAggregateBuilder
    Set Builder = New App_EnrollmentAggregateBuilder
    Set Load = Builder.Build(RowsOfColumnDtos)
End Function

Private Function CreateEnrollmentRepository() As Inf_EnrollmentRepository
    Set CreateEnrollmentRepository = New Inf_EnrollmentRepository
    CreateEnrollmentRepository.Initialize New Dom_SchoolYearCalculator, CreateFilePathResolver
End Function

Private Function CreateFilePathResolver() As Inf_FilePathResolver
    Set CreateFilePathResolver = New Inf_FilePathResolver
    CreateFilePathResolver.Initialize New Inf_WorkbookPathProvider
End Function
