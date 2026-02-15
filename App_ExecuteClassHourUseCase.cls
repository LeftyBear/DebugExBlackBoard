VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteClassHourUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.UseCase"
Option Explicit

Private Type Member
    Resolver As App_FilePathResolver
    Repository As Inf_ClassHourRepository
End Type
Private This As Member

Public Sub Initialize(ByVal Resolver As App_FilePathResolver, ByVal Repository As Inf_ClassHourRepository)
    Set This.Resolver = Resolver
    Set This.Repository = Repository
End Sub

Public Function Execute(ByVal TargetDate As Date, Optional ByVal KeywordText As String = VBA.vbNullString) As App_ExecuteClassHourResult
    Dim FilePath As String
    FilePath = This.Resolver.ClassHourFileByYear(TargetDate)
    Dim Raw As VBA.Collection
    Set Raw = This.Repository.ReadAll(FilePath)
    Dim List As Dom_ClassHourList
    Set List = New Dom_ClassHourList
    List.Initialize Raw
    Dim Summary As Dom_ClassHourSummary
    Set Summary = New Dom_ClassHourSummary
    Summary.Inject List.Items
    Dim Result As App_ExecuteClassHourResult
    Set Result = New App_ExecuteClassHourResult
    Result.Initialize List, Summary
    Set Execute = Result
End Function


