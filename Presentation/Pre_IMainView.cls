VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_IMainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Interface")
Option Explicit

Public Sub ShowLoading()
End Sub

Public Sub HideLoading()
End Sub

Public Sub ShowSuccess(ByVal Message As String)
End Sub

Public Sub NotifyBusinessError(ByVal Message As String)
End Sub

Public Sub NotifySystemError()
End Sub

Public Sub ShowSchedule(ByVal ViewModels As App_ScheduleReadModels)
End Sub

Public Sub ShowDailyPeriod(ByRef ViewTable() As Variant)
End Sub

Public Sub RenderClassHourPlan(ByRef ViewTable() As Variant)
End Sub

Public Sub RenderClassHourExecution(ByRef ViewTable() As Variant)
End Sub

Public Sub RenderTimeTablePlan(ByRef ViewTable() As Variant)
End Sub

Public Sub RenderTimeTableExecution(ByRef ViewTable() As Variant)
End Sub

Public Sub RenderEnrollment(ByRef ViewTable() As Variant)
End Sub
