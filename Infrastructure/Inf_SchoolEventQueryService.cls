VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolEventQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_ISchoolEventQueryService
Private Type Member
    Persistence As Inf_SchoolEventPersistence
    Rows        As Inf_SchoolEventRows
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_SchoolEventPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Rows Is Nothing Then Exit Sub
    Set This.Rows = New Inf_SchoolEventRows
    Dim Rows As Inf_SchoolEventRows
    Set Rows = This.Persistence.ReadAll
    BuildRow Rows
End Sub

Private Sub BuildRow(ByVal Rows As Inf_SchoolEventRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SchoolEventRow
        Set Row = Rows.Item(i)
        This.Rows.Add Row
    Next
End Sub

Private Function App_ISchoolEventQueryService_GetAll() As App_SchoolEventReadModels
    EnsureLoaded
    Set App_ISchoolEventQueryService_GetAll = Inf_SchoolEventModelFactory.Create(This.Rows)
End Function
