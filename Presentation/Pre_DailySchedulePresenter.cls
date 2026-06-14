VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_DailySchedulePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailySchedulePresenter

Private Function App_IDailySchedulePresenter_CreateViewModel(ByVal ReadModels As App_ScheduleReadModels) As Pre_ScheduleViewModels
    Dim Result As Pre_ScheduleViewModels
    Set Result = New Pre_ScheduleViewModels
    Dim i As Long
    For i = 1 To ReadModels.Count
        Dim ViewModel As Pre_ScheduleViewModel
        Set ViewModel = New Pre_ScheduleViewModel
        ViewModel.Name = ReadModels.Item(i).Name
        ViewModel.Value = ReadModels.Item(i).Value
        Result.Add ViewModel
    Next
    Set App_IDailySchedulePresenter_CreateViewModel = Result
End Function
