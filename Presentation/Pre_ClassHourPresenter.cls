VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_ClassHourPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyClassHourPresenter

Private Function App_IDailyClassHourPresenter_CreateViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = Format(ClassHourReadModels, SubjectReadModels, Totals, GradeCount)
    Set App_IDailyClassHourPresenter_CreateViewModel = Result
End Function

Private Function Format(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildBody(ClassHourReadModels, SubjectReadModels, Totals, GradeCount)
    Format = Pre_ArrayUtility.MergeArray(Header, Body)
End Function

Private Function BuildHeader(ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount + 1)
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(1, Grade + 1) = Grade & "”N"
    Next
    BuildHeader = Result
End Function

Private Function BuildBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Variant()
    Dim BottomIndex As Long
    BottomIndex = SubjectReadModels.Count + 1
    Dim Result() As Variant
    ReDim Result(1 To BottomIndex, 1 To GradeCount + 1)
    Dim i As Long
    For i = 1 To ClassHourReadModels.ClassHourCount
        Dim Model As App_ClassHourReadModel
        Set Model = ClassHourReadModels.ClassHourItem(i)
        Dim SubjectIndex As Long
        SubjectIndex = GetSubjectIndex(SubjectReadModels, Model.Subject)
        Result(SubjectIndex, Model.Grade + 1) = Model.Value
    Next
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(BottomIndex, Grade + 1) = Totals.Item(Grade)
    Next
    BuildBody = Result
End Function

Private Function GetSubjectIndex(ByVal SubjectReadModels As App_SubjectReadModels, ByVal SubjectName As String) As Long
    Dim i As Long
    For i = 1 To SubjectReadModels.Count
        If SubjectReadModels.Item(i).Name = SubjectName Then GetSubjectIndex = i: Exit For
    Next
End Function
