VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateSchoolStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Builder     As App_ConfigFilePathBuilder
    Repository  As Dom_IConfigRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_ConfigFilePathBuilder, ByVal Repository As Dom_IConfigRepository)
    Set This.Builder = Builder
    Set This.Repository = Repository
End Sub

Public Function Execute() As Dom_SchoolStructure
    Dim UpperValueFilePath As String
    UpperValueFilePath = This.Builder.Builed(UpperValues)
    If VBA.Dir(UpperValueFilePath) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "File path is Invalid."
    Dim UpperValueRows() As String
    UpperValueRows = This.Repository.Load(UpperValueFilePath)
    If App_Util_Array.IsEmpty(UpperValueRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "RawCSVRows is empty."
    Dim UpperValueFactory As New Dom_UpperValueFactory
    Dim UpperValue As Dom_UpperValue
    Set UpperValue = UpperValueFactory.Create(UpperValueRows)
    Dim SpecialStreamFilePath As String
    SpecialStreamFilePath = This.Builder.Builed(SpecialStream)
    If VBA.Dir(SpecialStreamFilePath) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "File path is Invalid."
    Dim SpecialStreamRows() As String
    SpecialStreamRows = This.Repository.Load(SpecialStreamFilePath)
    If App_Util_Array.IsEmpty(SpecialStreamRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "RawCSVRows is empty."
    Dim SpecialStreamCatalogFactory As New Dom_SpecialStreamCatalogFactory
    Dim SpecialStreamCatalog As Dom_SpecialStreamCatalog
    Set SpecialStreamCatalog = SpecialStreamCatalogFactory.Create(SpecialStreamRows)
    Dim SubjectFilePath As String
    SubjectFilePath = This.Builder.Builed(Subject)
    If VBA.Dir(SubjectFilePath) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "File path is Invalid."
    Dim SubjectRows() As String
    SubjectRows = This.Repository.Load(Subject)
    If App_Util_Array.IsEmpty(SubjectRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "RawCSVRows is empty."
    Dim SubjectCatalogFactory As New Dom_SubjectCatalogFactory
    Dim SubjectCatalog As Dom_SubjectCatalog
    Set SubjectCatalog = SubjectCatalogFactory.Create(SubjectRows)
    Dim Result As Dom_SchoolStructure
    Set Result = New Dom_SchoolStructure
    Result.Initialize UpperValue, SpecialStreamCatalog, SubjectCatalog
    Set Execute = Result
End Function
