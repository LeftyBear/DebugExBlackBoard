VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.QueryService"
Option Explicit

Private Type Member
    Persistence As Inf_EnrollmentPersistence
    DateIndex   As Scripting.Dictionary
End Type

Private This As Member
Private Const DateColumns As Long = 1

Public Sub Initialize(ByVal Persistence As Inf_CSVPersistence)
    Set This.Persistence = Persistence
    Set This.DateIndex = New Scripting.Dictionary
End Sub

Private Sub BuildDateIndex(ByVal TargetDate As Date)
    Dim SchoolYear As Long
    SchoolYear = Inf_DateUtility.GetSchoolYear(TargetDate)
    Dim Rows As Inf_EnrollmentRows
    Set Rows = This.Persistence.ReadAll(SchoolYear)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_EnrollmentRow
        Set Row = Rows.Item(i)
        If Row.NormDate <> vbNullString Then
            Dim Key As String
            Key = Row.NormDate
            If Not This.DateIndex.Exists(Key) Then
                This.DateIndex.Add Key, New VBA.Collection
            End If
            Set This.DateIndex.Item(Key) = Row: Exit For
        End If
    Next
End Sub

Private Function App_IEnrollmentQueryService_GetByDate(ByVal TargetDate As Date) As Inf_EnrollmentRows
    Dim Key As String
    Key = CStr(TargetDate)
    If This.DateIndex Is Nothing Then
        BuildDateIndex TargetDate
    Else
        If Not This.DateIndex.Exists(Key) Then BuildDateIndex TargetDate
    End If
    If Not This.DateIndex.Exists(Key) Then Err.Raise InfErrNotExistsItem, "Inf_EnrollmentQueryService", "Not exists item.  ( Key: " & Key & " )"
    Set App_IEnrollmentQueryService_GetByDate = This.DateIndex.Item(Key)
End Function

Private Function App_IEnrollmentQueryService_GetClassTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long, ByVal ClassNo As Long) As Long

End Function

Private Function App_IEnrollmentQueryService_GetFemaleTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long) As Long

End Function

Private Function App_IEnrollmentQueryService_GetMaleTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long) As Long

End Function
