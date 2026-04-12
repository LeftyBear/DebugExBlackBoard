VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Builder     As App_EntityFilePathBuilder
    Reader      As Dom_IClassHourReader
    Resolver    As Dom_ClassHourColumnResolver
    Factory     As Dom_ClassHourYearAggFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_EntityFilePathBuilder, ByVal Reader As Dom_IClassHourReader)
    Set This.Builder = Builder
    Set This.Reader = Reader
    Set This.Resolver = New Dom_ClassHourColumnResolver
    Set This.Factory = New Dom_ClassHourYearAggFactory
End Sub

Public Function Execute(ByVal SelectedDate As Date, ByVal Structure As Dom_SchoolStructure) As Dom_ClassHourYearAggregate
    With New Dom_SchoolYear
        Dim SchoolYear As Dom_SchoolYear
        Set SchoolYear = .Create(Dom_Util_Date.GetSchoolYear(SelectedDate))
    End With
    Dim FilePath As String
    FilePath = This.Builder.Builed(ClassHour, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateClassHour", "File path is invalid."
    Dim RawCSVRows() As String
    RawCSVRows = This.Reader.Load(FilePath)
    If App_Util_Array.IsEmpty(RawCSVRows) Then Err.Raise AppErrEmptyArray, "App_AggregateClassHour", "RawCSVRows is empty."
    This.Resolver.AddOn Structure
    Dim ColumnGroup As VBA.Collection
    Set ColumnGroup = This.Resolver.Resolve(RawCSVRows)
    If ColumnGroup.Count = 0 Then Err.Raise AppErrNothingItem, "App_AggregateClassHour", "ColumnGroup item is nothing."
    Set Execute = This.Factory.Create(ColumnGroup)
End Function
