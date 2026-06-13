VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_DailyPeriodPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presenter.ViewModel")
Option Explicit

Private Type Member
    View        As Pre_IMainView
    Formatter   As App_PeriodFormatter
End Type

Private This As Member

Friend Sub Inject(ByVal View As Pre_IMainView, ByVal Formatter As App_PeriodFormatter)
    Set This.View = View
    Set This.Formatter = Formatter
End Sub

Public Sub Present(ByVal ClassHourModels As App_ClassHourReadModels, ByVal MainStreamModel As App_MainStreamReadModel)
    Dim Table() As Variant
    Table = This.Formatter.Format(ExeKeyCode, ClassHourModels, MainStreamModel)
    This.View.ShowDailyPeriod Table
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
