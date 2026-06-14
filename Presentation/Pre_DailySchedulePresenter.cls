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

Public Function CreateViewModel(ByVal ViewModels As App_ScheduleReadModels) As App_ScheduleReadModels
    Dim ViewModels As App_ScheduleReadModels
    Set CreateViewModel = ViewModels
End Function
