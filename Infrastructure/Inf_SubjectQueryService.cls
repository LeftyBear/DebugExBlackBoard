VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.QueryService"
Option Explicit
Implements App_ISubjectQueryService
Private Type Member
    Persistence As Inf_SubjectPersistence
    NameIndex   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_SubjectPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If This.NameIndex Is Nothing Then
        Set This.NameIndex = New Scripting.Dictionary
        BuildNameIndex
    End If
End Sub

Private Sub BuildNameIndex()
    Dim RawRows As Inf_RawRows
    Set RawRows = This.Persistence.ReadAll
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim Row As Inf_SubjectRow
        Set Row = New Inf_SubjectRow
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            If RawRows.GetField(1, C) = jpnKyoukaMei Then
                Row.Name = RawRows.GetField(R, C)
            ElseIf Rows.GetField(1, C) = jpnTaisyoGakunen Then
                Row.TargetGrade = RawRows.GetField(R, C)
            ElseIf Rows.GetField(1, C) = jpnKigou Then
                Row.Mark = RawRows.GetField(R, C)
            End If
        Next
        This.NameIndex.Add Row.Name, Row
    Next
End Sub

Private Function App_ISubjectQueryService_GetByName(ByVal Name As String) As App_SubjectReadModel
    EnsureLoaded
    If This.NameIndex.Exists(Name) Then Exit Function
    Dim Row As Inf_SubjectRow
    Set Row = This.NameIndex.Item(Name)
    Dim Factory As Inf_SubjectModelFactory
    Set Factory = New Inf_SubjectModelFactory
    Set App_ISubjectQueryService_GetByName = Factory.Create(Row)
End Function
