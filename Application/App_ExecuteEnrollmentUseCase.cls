VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteEnrollmentUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.UseCase"
Option Explicit

Private Type Member
    Resolver As App_FilePathResolver
    Repository As Inf_EnrollmentRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Resolver As App_FilePathResolver, ByVal Repository As Inf_EnrollmentRepository)
    Set This.Resolver = Resolver
    Set This.Repository = Repository
End Sub

Public Function Execute(ByVal TargetDate As Date, Optional ByVal KeywordText As String = VBA.vbNullString) As App_ExecuteEnrollmentResult
    Dim FilePath As String
    FilePath = This.Resolver.EnrollmentFileByYear(TargetDate)
    Dim Raw As VBA.Collection
    Set Raw = This.Repository.ReadAll(FilePath)
    Dim List As Dom_EnrollmentList
    Set List = New Dom_EnrollmentList
    List.Initialize Raw
    If 0 < VBA.Len(KeywordText) Then
        Dim Keywords As VBA.Collection
        Set Keywords = Util_Function.ParseSearchKeywords(KeywordText)
        Set List = List.SearchRemarksByKeywords(Keywords)
    End If
    Dim Summary As Dom_EnrollmentSummary
    Set Summary = New Dom_EnrollmentSummary
    Summary.Inject List.Items
    Dim Result As App_ExecuteEnrollmentResult
    Set Result = New App_ExecuteEnrollmentResult
    Result.Initialize List, Summary
    Set Execute = Result
End Function

