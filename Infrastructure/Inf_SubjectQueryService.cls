VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_ISubjectQueryService
Private Type Member
    Persistence As Inf_SubjectPersistence
    Rows        As Inf_SubjectRows
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_SubjectPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Rows Is Nothing Then Exit Sub
    Set This.Rows = This.Persistence.ReadAll
End Sub

Private Function App_ISubjectQueryService_GetAll() As App_SubjectReadModels
    EnsureLoaded
    Set App_ISubjectQueryService_GetAll = Inf_SubjectModelFactory.Create(This.Rows)
End Function
