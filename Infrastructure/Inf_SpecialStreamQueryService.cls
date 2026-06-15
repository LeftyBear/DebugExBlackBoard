VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_ISpecialStreamQueryService
Private Type Member
    Persistence As Inf_SpecialStreamPersistence
    Rows        As Inf_SpecialStreamRows
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_SpecialStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Rows Is Nothing Then Exit Sub
    Dim Rows As Inf_SpecialStreamRows
    Set Rows = This.Persistence.ReadAll
    BuildRow Rows
End Sub

Private Sub BuildRow(ByVal Rows As Inf_SpecialStreamRows)
    Dim i As Long
    For i = 1 To Rows.Count
        This.Rows.Add Rows.Item(i)
    Next
End Sub

Private Function App_ISpecialStreamQueryService_GetAll() As App_SpecialStreamReadModels
    EnsureLoaded
    Set App_ISpecialStreamQueryService_GetAll = Inf_SpecialStreamModelFactory.Create(This.Rows)
End Function

Private Sub App_ISpecialStreamQueryService_Rebuild()
    Dim Rows As Inf_SpecialStreamRows
    Set Rows = This.Persistence.ReadAll
    BuildRow Rows
End Sub
