VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.QueryService"
Option Explicit
Implements App_ISpecialStreamQueryService
Private Type Member
    Persistence As Inf_SpecialStreamPersistence
    Rows        As Inf_SpecialStreamRows
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_SpecialStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Rows Is Nothing Then Exit Sub
    Set This.Rows = New Inf_SpecialStreamRows
    Dim RawRows As Inf_RawRows
    Set RawRows = This.Persistence.ReadAll
    BuildRow RawRows
End Sub

Private Sub BuildRow(ByVal RawRows As Inf_RawRows)
    Dim Factory As Inf_SpecialStreamRowFactory
    Set Factory = New Inf_SpecialStreamRowFactory
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            Dim Row As Inf_SpecialStreamRow
            Set Row = Factory.Create(RawRows.GetField(1, C), RawRows.GetField(R, C))
            This.Rows.Add Row
        Next
    Next
End Sub

Private Function App_ISpecialStreamQueryService_GetAll() As App_SpecialStreamReadModels
    EnsureLoaded
    Dim Factory As New Inf_SpecialStreamModelFactory
    Set App_ISpecialStreamQueryService_GetAll = Factory.Create(This.Rows)
End Function
