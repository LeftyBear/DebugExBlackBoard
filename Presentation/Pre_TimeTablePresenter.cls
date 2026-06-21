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

Private Function App_IDailyTimeTablePresenter_CreateViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal PeriodCount As Long, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = Format(ClassHourReadModels, Totals, PeriodCount, GradeCount)
    Set App_IDailyTimeTablePresenter_CreateViewModel = Result
End Function

Private Function Format(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal PeriodCount As Long, ByVal GradeCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildBody(ClassHourReadModels, Totals, PeriodCount, GradeCount)
    Format = Pre_ArrayUtility.MergeArray(Header, Body)
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

Private Function BuildBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Totals As App_PeriodTotals, ByVal PeriodCount As Long, ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To PeriodCount + 1, 1 To GradeCount)
    Dim i As Long
    For i = 1 To ClassHourReadModels.TimeTableCount
        Dim Model As App_TimeTableReadModel
        Set Model = ClassHourReadModels.TimeTableItem(i)
        Result(Model.Period, Model.Grade) = Model.Mark
    Next
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(PeriodCount + 1, Grade) = Totals.Item(Grade)
    Next
    BuildBody = Result
End Function
