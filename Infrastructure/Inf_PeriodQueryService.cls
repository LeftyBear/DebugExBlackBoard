VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_PeriodQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_IPeriodQueryService
Private Type Member
    Persistence As Inf_PeriodPersistence
    Row         As Inf_PeriodRow
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_PeriodPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Row Is Nothing Then Exit Sub
    Set This.Row = This.Persistence.ReadAll
End Sub

Private Function App_IPeriodQueryService_GetModel() As App_PeriodReadModel
    EnsureLoaded
    Dim Factory As New Inf_PeriodModelFactory
    Set App_IPeriodQueryService_GetModel = Factory.Create(This.Row)
End Function
