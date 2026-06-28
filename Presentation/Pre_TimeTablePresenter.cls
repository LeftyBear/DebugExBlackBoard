VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_TimeTablePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyTimeTablePresenter

Private Function FormatPlan(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal GradeCount As Long, ByVal PeriodCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildPlanBody(ClassHourReadModels, Totals, GradeCount, PeriodCount)
    FormatPlan = Pre_ArrayUtility.MergeVertically(Header, Body)
End Function

Private Function FormatExe(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal GradeCount As Long, ByVal PeriodCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildExeBody(ClassHourReadModels, Totals, GradeCount, PeriodCount)
    FormatExe = Pre_ArrayUtility.MergeVertically(Header, Body)
End Function

Private Function BuildHeader(ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount)
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(1, Grade) = Grade & "”N"
    Next
    BuildHeader = Result
End Function

Private Function BuildPlanBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal GradeCount As Long, ByVal PeriodCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To PeriodCount + 1, 1 To GradeCount)
    Dim i As Long
    For i = 1 To ClassHourReadModels.TimeTablePlanCount
        Dim Model As App_TimeTableReadModel
        Set Model = ClassHourReadModels.GetTimeTablePlan(i)
        Result(Model.Period, Model.Grade) = Model.Mark
    Next
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(PeriodCount + 1, Grade) = Totals.Item(Grade)
    Next
    BuildPlanBody = Result
End Function

Private Function BuildExeBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal GradeCount As Long, ByVal PeriodCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To PeriodCount + 1, 1 To GradeCount)
    Dim i As Long
    For i = 1 To ClassHourReadModels.TimeTableExeCount
        Dim Model As App_TimeTableReadModel
        Set Model = ClassHourReadModels.GetTimeTableExe(i)
        Result(Model.Period, Model.Grade) = Model.Mark
    Next
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(PeriodCount + 1, Grade) = Totals.Item(Grade)
    Next
    BuildExeBody = Result
End Function

Private Function App_IDailyTimeTablePresenter_CreateExeViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal GradeCount As Long, ByVal PeriodCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = FormatExe(ClassHourReadModels, Totals, GradeCount, PeriodCount)
    Set App_IDailyTimeTablePresenter_CreateExeViewModel = Result
End Function

Private Function App_IDailyTimeTablePresenter_CreatePlanViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal GradeCount As Long, ByVal PeriodCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = FormatPlan(ClassHourReadModels, Totals, GradeCount, PeriodCount)
    Set App_IDailyTimeTablePresenter_CreatePlanViewModel = Result
End Function
