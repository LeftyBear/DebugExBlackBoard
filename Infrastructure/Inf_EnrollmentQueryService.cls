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
Implements App_IEnrollmentQueryService
Private Type Member
    Persistence As Inf_EnrollmentPersistence
    DateIndex   As Scripting.Dictionary
End Type

Private This As Member
Private Const DateColumns As Long = 1

Public Sub Initialize(ByVal Persistence As Inf_EnrollmentPersistence)
    Set This.Persistence = Persistence
    Set This.DateIndex = New Scripting.Dictionary
End Sub

Private Sub EnsureLoaded(ByVal TargetDate As Date)
    If This.DateIndex Is Nothing Then
        BuildDateIndex TargetDate
    Else
        If Not This.DateIndex.Exists(CStr(TargetDate)) Then
            BuildDateIndex TargetDate
        End If
    End If
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
            Dim DailyRows As Inf_EnrollmentRows
            Set DailyRows = New Inf_EnrollmentRows
            This.DateIndex.Add Row.NormDate, DailyRows
        End If
        DailyRows.Add Row
    Next
End Sub

Private Function App_IEnrollmentQueryService_GetByDate(ByVal TargetDate As Date) As App_EnrollmentReadModels
    EnsureLoaded TargetDate
    If This.DateIndex.Exists(CStr(TargetDate)) Then Exit Function
    Dim DailyRows As Inf_EnrollmentRows
    Set DailyRows = This.DateIndex.Item(CStr(TargetDate))
    Dim Factory As New Inf_DailyEnrollmentModelFactory
    Set App_IEnrollmentQueryService_GetByDate = Factory.Create(DailyRows)
End Function
