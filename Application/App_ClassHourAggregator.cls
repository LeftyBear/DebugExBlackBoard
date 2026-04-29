VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourAggregator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Reader      As App_IClassHourReader
    Builder     As App_EntityFilePathBuilder
    ColResolver As App_ClassHourColumnResolver
    RowResolver As App_ClassHourRowResolver
    Factory     As App_ClassHourYearAggFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As App_IClassHourReader, ByVal Builder As App_EntityFilePathBuilder, ByVal ColResolver As App_ClassHourColumnResolver, ByVal RowResolver As App_ClassHourRowResolver, ByVal Factory As App_ClassHourYearAggFactory)
    Set This.Reader = Reader
    Set This.Builder = Builder
    Set This.ColResolver = ColResolver
    Set This.RowResolver = RowResolver
    Set This.Factory = Factory
End Sub

Public Function Execute(ByVal SchoolYear As Dom_SchoolYear, ByVal Structure As Dom_SchoolStructure) As Dom_IClassHourRepository
    Dim FilePath As String
    FilePath = This.Builder.Builed(ClassHour, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim RawRows() As String
    RawRows = This.Reader.Load(FilePath)
    If App_ArrayUtility.IsEmpty(RawRows) Then Err.Raise AppErrEmptyArray, "App_ClassHourAggregator", "RawRows is empty."
    Dim HeaderRow() As String
    HeaderRow = App_ArrayUtility.GetHeaderRow(RawRows)
    Dim MetaMap As Scripting.Dictionary
    Set MetaMap = This.ColResolver.Resolve(HeaderRow)
    Dim Rows As VBA.Collection
    Set Rows = This.RowResolver.Resolve(RawRows, MetaMap)
    If Rows.Count = 0 Then Err.Raise AppErrNothingItem, "App_ClassHourAggregator", "Rows item is nothing."
    Dim Repository As Dom_IClassHourRepository
    Set Repository = New Inf_ClassHourRepository
    Dim YearAggregate As Dom_ClassHourYearAggregate
    Set YearAggregate = This.Factory.Create(Rows)
    Dim i As Long
    For i = 1 To YearAggregate.Count
        Dim Aggregate As Dom_ClassHourAggregate
        Set Aggregate = YearAggregate.Item(i)
        Repository.Add Aggregate
    Next
    Set Execute = Repository
End Function
