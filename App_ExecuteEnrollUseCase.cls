VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteEnrollUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.UseCase"
Option Explicit

Private Type Member
    Repository As Inf_EnrollRepository
End Type
Private This As Member

Public Sub Initialize(ByVal Repository As Inf_EnrollRepository)
    Set This.Repository = Repository
End Sub

Public Function Execute(Optional ByVal KeywordText As String = VBA.vbNullString) As App_ExecuteEnrollResult
    Dim Raw As VBA.Collection
    Set Raw = This.Repository.ReadAll
    Dim List As Dom_EnrollMemberList
    Set List = New Dom_EnrollMemberList
    List.Initialize Raw
    If 0 < VBA.Len(KeywordText) Then
        Dim Keywords As VBA.Collection
        Set Keywords = Util_Function.ParseSearchKeywords(KeywordText)
        Set List = List.SearchRemarksByKeywords(Keywords)
    End If
    Dim Summary As Dom_EnrollSummary
    Set Summary = New Dom_EnrollSummary
    Summary.Inject List.Items
    Dim Result As App_ExecuteEnrollResult
    Set Result = New App_ExecuteEnrollResult
    Result.Initialize List, Summary
    Set Execute = Result
End Function

