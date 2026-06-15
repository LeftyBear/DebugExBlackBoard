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

Public Function App_IDailyPeriodPresenter_CreateViewModel(ByVal Periods As App_Periods, ByVal GradeCount As Long) As Pre_PeriodViewModel
    Dim Result As Pre_PeriodViewModel
    Set Result = New Pre_PeriodViewModel
    Dim Formatter As Pre_PeriodFormatter
    Set Formatter = New Pre_PeriodFormatter
    Result.Table = Formatter.Format(Periods, GradeCount)
    Set App_IDailyPeriodPresenter_CreateViewModel = Result
End Function
