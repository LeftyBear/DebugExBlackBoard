VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_IMainStreamQueryService
Private Type Member
    Persistence As Inf_MainStreamPersistence
    Rows        As Inf_MainStreamRows
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_MainStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Rows Is Nothing Then Exit Sub
    Set This.Rows = This.Persistence.ReadAll
End Sub

Private Function App_IMainStreamQueryService_GetAll() As App_MainStreamReadModel
    EnsureLoaded
    Set App_IMainStreamQueryService_GetAll = Inf_MainStreamModelFactory.Create(This.Rows)
End Function

Private Sub App_IMainStreamQueryService_Rebuild()
    Set This.Rows = This.Persistence.ReadAll
End Sub
