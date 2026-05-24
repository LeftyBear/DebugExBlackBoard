VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.QueryService"
Option Explicit
Implements App_IMainStreamQueryService
Private Type Member
    Persistence As Inf_MainStreamPersistence
    Rows        As Inf_MainStreamRows
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_MainStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Rows Is Nothing Then Exit Sub
    Set This.Rows = New Inf_MainStreamRows
    Dim RawRows As Inf_RawRows
    Set RawRows = This.Persistence.ReadAll
    BuildRows RawRows
End Sub

Private Sub BuildRows(ByVal RawRows As Inf_RawRows)
    Dim Factory As Inf_MainStreamRowFactory
    Set Factory = New Inf_MainStreamRowFactory
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            Dim Row As Inf_MainStreamRow
            Set Row = Factory.Create(RawRows.GetField(1, C), RawRows.GetField(R, C))
            This.Rows.Add Row
        Next
    Next
End Sub

Private Function App_IMainStreamQueryService_GetModel() As App_MainStreamReadModel
    EnsureLoaded
    Dim Factory As New Inf_MainStreamModelFactory
    Set App_IMainStreamQueryService_GetModel = Factory.Create(This.Rows)
End Function
