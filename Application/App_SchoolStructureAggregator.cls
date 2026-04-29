VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SchoolStructureAggregator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Reader                      As App_IConfigReader
    Builder                     As App_ConfigFilePathBuilder
    UpperValueFactory           As App_UpperValueFactory
    SpecialStreamCatalogFactory As App_SpecialStreamCatalogFactory
    SubjectCatalogFactory       As App_SubjectCatalogFactory
    SchoolEventCatalogFactory   As App_SchoolEventCatalogFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As App_IConfigReader, ByVal Builder As App_ConfigFilePathBuilder, ByVal UpperValueFactory As App_UpperValueFactory, ByVal SpecialStreamCatalogFactory As App_SpecialStreamCatalogFactory, ByVal SubjectCatalogFactory As App_SubjectCatalogFactory, ByVal SchoolEventCatalogFactory As App_SchoolEventCatalogFactory)
    Set This.Reader = Reader
    Set This.Builder = Builder
    Set This.UpperValueFactory = UpperValueFactory
    Set This.SpecialStreamCatalogFactory = SpecialStreamCatalogFactory
    Set This.SubjectCatalogFactory = SubjectCatalogFactory
    Set This.SchoolEventCatalogFactory = SchoolEventCatalogFactory
End Sub

Public Function Execute() As Dom_SchoolStructure
    'UpperValue-----------------------------------------------------------------------------------
    Dim UpperValueFilePath As String
    UpperValueFilePath = This.Builder.Builed(UpperValues)
    If VBA.Len(VBA.Dir(UpperValueFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_SchoolStructureAggregator", "UpperValue file path is Invalid."
    Dim UpperValueRows() As String
    UpperValueRows = This.Reader.Load(UpperValueFilePath)
    If App_ArrayUtility.IsEmpty(UpperValueRows) Then Err.Raise AppErrEmptyArray, "App_SchoolStructureAggregator", "UpperValue RawRows is empty."
    Dim UpperValue As Dom_UpperValue
    Set UpperValue = This.UpperValueFactory.Create(UpperValueRows)
    'SpecialStreamCatalog-------------------------------------------------------------------------
    Dim SpecialStreamFilePath As String
    SpecialStreamFilePath = This.Builder.Builed(SpecialStream)
    If VBA.Len(VBA.Dir(SpecialStreamFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_SchoolStructureAggregator", "SpecialStream file path is Invalid."
    Dim SpecialStreamRows() As String
    SpecialStreamRows = This.Reader.Load(SpecialStreamFilePath)
    If App_ArrayUtility.IsEmpty(SpecialStreamRows) Then Err.Raise AppErrEmptyArray, "App_SchoolStructureAggregator", "SpecialStream RawRows is empty."
    Dim SpecialStreamCatalog As Dom_SpecialStreamCatalog
    Set SpecialStreamCatalog = This.SpecialStreamCatalogFactory.Create(SpecialStreamRows)
    'SubjectCatalog-------------------------------------------------------------------------------
    Dim SubjectFilePath As String
    SubjectFilePath = This.Builder.Builed(Subject)
    If VBA.Len(VBA.Dir(SubjectFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_SchoolStructureAggregator", "Subject file path is Invalid."
    Dim SubjectRows() As String
    SubjectRows = This.Reader.Load(SubjectFilePath)
    If App_ArrayUtility.IsEmpty(SubjectRows) Then Err.Raise AppErrEmptyArray, "App_SchoolStructureAggregator", "Subject RawRows is empty."
    Dim SubjectCatalog As Dom_SubjectCatalog
    Set SubjectCatalog = This.SubjectCatalogFactory.Create(SubjectRows)
    'SchoolEventCatalog---------------------------------------------------------------------------
    Dim SchoolEventFilePath As String
    SchoolEventFilePath = This.Builder.Builed(Subject)
    If VBA.Len(VBA.Dir(SchoolEventFilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_SchoolStructureAggregator", "SchoolEvent file path is Invalid."
    Dim SchoolEventRows() As String
    SchoolEventRows = This.Reader.Load(SchoolEventFilePath)
    If App_ArrayUtility.IsEmpty(SchoolEventRows) Then Err.Raise AppErrEmptyArray, "App_SchoolStructureAggregator", "SchoolEvent RawRows is empty."
    Dim SchoolEventCatalog As Dom_SchoolEventCatalog
    Set SchoolEventCatalog = This.SchoolEventCatalogFactory.Create(SchoolEventRows)
    'SchoolEventStructure-------------------------------------------------------------------------
    Dim Result As Dom_SchoolStructure
    Set Result = New Dom_SchoolStructure
    Result.Initialize UpperValue, SpecialStreamCatalog, SubjectCatalog, SchoolEventCatalog
    Set Execute = Result
End Function
