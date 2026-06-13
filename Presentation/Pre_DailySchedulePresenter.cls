VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_DailySchedulePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presenter.ViewModel")
Option Explicit

Private Type Member
    View As Pre_IMainView
End Type

Private This As Member

Friend Sub Inject(ByVal View As Pre_IMainView)
    Set This.View = View
End Sub

Public Sub Present(ByVal ViewModels As App_ScheduleReadModels)
    This.View.ShowSchedule ViewModels
End Sub

Public Sub ShowSuccess(ByVal Message As String)
    This.View.ShowSuccess Message
End Sub

Public Sub NotifyBusinessError(ByVal Message As String)
    This.View.NotifyBusinessError Message
End Sub

Public Sub NotifySystemError()
    This.View.NotifySystemError
End Sub
