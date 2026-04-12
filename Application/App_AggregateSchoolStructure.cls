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
    Builder As App_ConfigFilePathBuilder
    Reader  As Dom_IConfigReader
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_ConfigFilePathBuilder, ByVal Reader As Dom_IConfigReader)
    Set This.Builder = Builder
    Set This.Reader = Reader
End Sub

Public Function Execute() As Dom_SchoolStructure
    'UpperValue-----------------------------------------------------------------------------------
    Dim UpperValueFilePath As String
    UpperValueFilePath = This.Builder.Builed(UpperValues)
    If VBA.Len(VBA.Dir(UpperValueFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "UpperValue file path is Invalid."
    Dim UpperValueRows() As String
    UpperValueRows = This.Reader.Load(UpperValueFilePath)
    If App_Util_Array.IsEmpty(UpperValueRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "UpperValue RawCSVRows is empty."
    '@Ignore SelfAssignedDeclaration
    Dim UpperValueFactory As New Dom_UpperValueFactory
    Dim UpperValue As Dom_UpperValue
    Set UpperValue = UpperValueFactory.Create(UpperValueRows)
    'SpecialStreamCatalog-------------------------------------------------------------------------
    Dim SpecialStreamFilePath As String
    SpecialStreamFilePath = This.Builder.Builed(SpecialStream)
    If VBA.Len(VBA.Dir(SpecialStreamFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "SpecialStream file path is Invalid."
    Dim SpecialStreamRows() As String
    SpecialStreamRows = This.Reader.Load(SpecialStreamFilePath)
    If App_Util_Array.IsEmpty(SpecialStreamRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "SpecialStream RawCSVRows is empty."
    '@Ignore SelfAssignedDeclaration
    Dim SpecialStreamCatalogFactory As New Dom_SpecialStreamCatalogFactory
    Dim SpecialStreamCatalog As Dom_SpecialStreamCatalog
    Set SpecialStreamCatalog = SpecialStreamCatalogFactory.Create(SpecialStreamRows)
    'SubjectCatalog-------------------------------------------------------------------------------
    Dim SubjectFilePath As String
    SubjectFilePath = This.Builder.Builed(Subject)
    If VBA.Len(VBA.Dir(SubjectFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "Subject file path is Invalid."
    Dim SubjectRows() As String
    SubjectRows = This.Reader.Load(SubjectFilePath)
    If App_Util_Array.IsEmpty(SubjectRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "Subject RawCSVRows is empty."
    '@Ignore SelfAssignedDeclaration
    Dim SubjectCatalogFactory As New Dom_SubjectCatalogFactory
    Dim SubjectCatalog As Dom_SubjectCatalog
    Set SubjectCatalog = SubjectCatalogFactory.Create(SubjectRows)
    'SchoolEventCatalog---------------------------------------------------------------------------
    Dim SchoolEventFilePath As String
    SchoolEventFilePath = This.Builder.Builed(Subject)
    If VBA.Len(VBA.Dir(SchoolEventFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateSchoolStructure", "SchoolEvent file path is Invalid."
    Dim SchoolEventRows() As String
    SchoolEventRows = This.Reader.Load(SchoolEventFilePath)
    If App_Util_Array.IsEmpty(SchoolEventRows) Then Err.Raise AppErrEmptyArray, "App_AggregateSchoolStructure", "SchoolEvent RawCSVRows is empty."
    '@Ignore SelfAssignedDeclaration
    Dim SchoolEventCatalogFactory As New Dom_SchoolEventCatalogFactory
    Dim SchoolEventCatalog As Dom_SchoolEventCatalog
    Set SchoolEventCatalog = SchoolEventCatalogFactory.Create(SchoolEventRows)
    'SchoolEventStructure-------------------------------------------------------------------------
    Dim Result As Dom_SchoolStructure
    Set Result = New Dom_SchoolStructure
    Result.Initialize UpperValue, SpecialStreamCatalog, SubjectCatalog, SchoolEventCatalog
    Set Execute = Result
End Function
