VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_EnrollmentPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyEnrollmentPresenter

Private Function App_IDailyEnrollmentPresenter_CreateViewModel(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainGradeCount As Long, ByVal MainClassNoCount As Long, ByVal SpecialClassNoCount As Long) As Pre_EnrollmentViewModel
    Dim Result As Pre_EnrollmentViewModel
    Set Result = New Pre_EnrollmentViewModel
    Dim Formatter As Pre_EnrollmentFormatter
    Set Formatter = New Pre_EnrollmentFormatter
    Result.Table = Formatter.Format(ReadModels, Totals, MainGradeCount, MainClassNoCount, SpecialClassNoCount)
    Set App_IDailyEnrollmentPresenter_CreateViewModel = Result
End Function
