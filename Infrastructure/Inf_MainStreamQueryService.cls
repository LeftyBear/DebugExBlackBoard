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
    Row         As Inf_MainStreamRow
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_MainStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Row Is Nothing Then Exit Sub
    Set This.Row = New Inf_MainStreamRow
    Dim RawRows As Inf_RawRows
    Set RawRows = This.Persistence.ReadAll
    BuildRow RawRows
End Sub

Private Sub BuildRow(ByVal RawRows As Inf_RawRows)
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            If RawRows.GetField(1, C) = jpnFutsuGakunenSu Then
                This.Row.UpperGrade = RawRows.GetField(R, C)
            ElseIf RawRows.GetField(1, C) = jpnFutsuGakkyuSu Then
                This.Row.UpperClassNo = RawRows.GetField(R, C)
            End If
        Next
    Next
End Sub

Private Function App_IMainStreamQueryService_GetModel() As App_MainStreamReadModel
    EnsureLoaded
    Dim Factory As New Inf_MainStreamModelFactory
    Set App_IMainStreamQueryService_GetModel = Factory.Create(This.Row)
End Function
