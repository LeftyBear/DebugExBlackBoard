VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit
Implements Dom_IEnrollmentRepository
Private Type Member
    Repository As Inf_CSVReadRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_CSVReadRepository)
    Set This.Repository = Repository
End Sub

Private Function Dom_IEnrollmentRepository_Load(ByVal SchoolYear As Long) As String()
    Dim RawRows() As String
    RawRows = This.Repository.Load(Dom_Entity.Enrollment, SchoolYear)
    Dom_IEnrollmentRepository_Load = RawRows
End Function
