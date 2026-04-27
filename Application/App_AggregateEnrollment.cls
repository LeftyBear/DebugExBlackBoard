VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Reader      As Dom_IEnrollmentReader
    Builder     As App_EntityFilePathBuilder
    ColResolver As App_EnrollmentColumnResolver
    RowResolver As App_EnrollmentRowResolver
    Factory     As App_EnrollmentYearAggFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As Dom_IEnrollmentReader, ByVal Builder As App_EntityFilePathBuilder, ByVal ColResolver As App_EnrollmentColumnResolver, ByVal RowResolver As App_EnrollmentRowResolver, ByVal Factory As App_EnrollmentYearAggFactory)
    Set This.Reader = Reader
    Set This.Builder = Builder
    Set This.ColResolver = ColResolver
    Set This.RowResolver = RowResolver
    Set This.Factory = Factory
End Sub

Public Function Execute(SchoolYear As Dom_SchoolYear) As App_EnrollmentRepository
    Dim FilePath As String
    FilePath = This.Builder.Builed(Enrollment, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim RawRows() As String
    RawRows = This.Reader.Load(FilePath)
    If App_ArrayUtility.IsEmpty(RawRows) Then Err.Raise AppErrEmptyArray, "App_AggregateEnrollment", "RawRows is empty."
    Dim HeaderRow() As String
    HeaderRow = App_ArrayUtility.GetHeaderRow(RawRows)
    Dim MetaMap As Scripting.Dictionary
    Set MetaMap = This.ColResolver.Resolve(HeaderRow)
    Dim Rows As VBA.Collection
    Set Rows = This.RowResolver.Resolve(RawRows, MetaMap)
    If Rows.Count = 0 Then Err.Raise AppErrNothingItem, "App_AggregateEnrollment", "Rows item is nothing."
    Dim Repository As App_EnrollmentRepository
    Set Repository = New App_EnrollmentRepository
    Dim YearAggregate As Dom_EnrollmentYearAggregate
    Set YearAggregate = This.Factory.Create(Rows)
    Dim i As Long
    For i = 1 To YearAggregate.Count
        Dim Aggregate As Dom_EnrollmentAggregate
        Set Aggregate = YearAggregate.Item(i)
        Repository.Add Aggregate
    Next
    Set Execute = Repository
End Function
