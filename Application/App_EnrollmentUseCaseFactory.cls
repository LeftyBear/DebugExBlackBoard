VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    CSV As Inf_CSVPersistence
End Type

Private This As Member

Public Sub Initialize(ByVal CSV As Inf_CSVPersistence)
    Set This.CSV = CSV
End Sub

Public Function Create() As App_LoadEnrollmentUseCase
    Dim Persistence As Inf_EnrollmentPersistence
    Set Persistence = New Inf_EnrollmentPersistence
    Persistence.Initialize This.CSV
    Dim Repository As Inf_EnrollmentRepository
    Set Repository = New Inf_EnrollmentRepository
    Repository.Initialize Persistence
    Dim Result As App_LoadEnrollmentUseCase
    Set Result = New App_LoadEnrollmentUseCase
    Result.Initialize Repository
    Set Create = Result
End Function
