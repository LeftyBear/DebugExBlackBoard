VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_PeriodPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyPeriodPresenter

Public Function App_IDailyPeriodPresenter_CreateViewModel(ByVal Periods As App_PeriodTotals, ByVal GradeCount As Long) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = Format(Periods, GradeCount)
    Set App_IDailyPeriodPresenter_CreateViewModel = Result
End Function

Private Function Format(ByVal Periods As App_PeriodTotals, ByVal GradeCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildBody(Periods, GradeCount)
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

Private Function BuildBody(ByVal Periods As App_PeriodTotals, ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount)
    If Not Periods Is Nothing Then
        Dim Grade As Long
        For Grade = 1 To GradeCount
            Result(1, Grade) = Periods.Item(Grade)
        Next
    End If
    BuildBody = Result
End Function
