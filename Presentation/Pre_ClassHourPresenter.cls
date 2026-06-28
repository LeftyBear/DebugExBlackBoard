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

Private Function FormatPlan(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildPlanBody(ClassHourReadModels, SubjectReadModels, Totals, GradeCount)
    FormatPlan = Pre_ArrayUtility.MergeVertically(Header, Body)
End Function

Private Function FormatExe(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildExeBody(ClassHourReadModels, SubjectReadModels, Totals, GradeCount)
    FormatExe = Pre_ArrayUtility.MergeVertically(Header, Body)
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

Private Function BuildPlanBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Variant()
    Dim BottomIndex As Long
    BottomIndex = SubjectReadModels.Count + 1
    Dim Result() As Variant
    ReDim Result(1 To BottomIndex, 1 To GradeCount + 1)
    Dim i As Long
    For i = 1 To ClassHourReadModels.ClassHourPlanCount
        Dim Model As App_ClassHourReadModel
        Set Model = ClassHourReadModels.GetClassHourPlan(i)
        Dim SubjectIndex As Long
        SubjectIndex = GetSubjectIndex(SubjectReadModels, Model.Subject)
        Result(SubjectIndex, Model.Grade + 1) = Model.Value
    Next
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(BottomIndex, Grade + 1) = Totals.Item(Grade)
    Next
    BuildPlanBody = Result
End Function

Private Function BuildExeBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Variant()
    Dim BottomIndex As Long
    BottomIndex = SubjectReadModels.Count + 1
    Dim Result() As Variant
    ReDim Result(1 To BottomIndex, 1 To GradeCount + 1)
    Dim i As Long
    For i = 1 To ClassHourReadModels.ClassHourExeCount
        Dim Model As App_ClassHourReadModel
        Set Model = ClassHourReadModels.GetClassHourExe(i)
        Dim SubjectIndex As Long
        SubjectIndex = GetSubjectIndex(SubjectReadModels, Model.Subject)
        Result(SubjectIndex, Model.Grade + 1) = Model.Value
    Next
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(BottomIndex, Grade + 1) = Totals.Item(Grade)
    Next
    BuildExeBody = Result
End Function

Private Function GetSubjectIndex(ByVal SubjectReadModels As App_SubjectReadModels, ByVal SubjectName As String) As Long
    Dim i As Long
    For i = 1 To SubjectReadModels.Count
        If SubjectReadModels.Item(i).Name = SubjectName Then GetSubjectIndex = i: Exit For
    Next
End Function

Private Function App_IDailyClassHourPresenter_CreateExeViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = FormatExe(ClassHourReadModels, SubjectReadModels, Totals, GradeCount)
    Set App_IDailyClassHourPresenter_CreateExeViewModel = Result
End Function

Private Function App_IDailyClassHourPresenter_CreatePlanViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = FormatPlan(ClassHourReadModels, SubjectReadModels, Totals, GradeCount)
    Set App_IDailyClassHourPresenter_CreatePlanViewModel = Result
End Function
