VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_IMainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Interface")
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

Public Sub Render(ByVal VM As Pre_ViewModel)
End Sub

Public Sub RenderEnrollment(ByRef Table() As Variant)
End Sub

Public Sub RenderClassHourPlan(ByRef Table() As Variant)
End Sub

Public Sub RenderClassHourExecution(ByRef Table() As Variant)
End Sub

Public Sub RenderTimeTablePlan(ByRef Table() As Variant)
End Sub

Public Sub RenderTimeTableExecution(ByRef Table() As Variant)
End Sub
