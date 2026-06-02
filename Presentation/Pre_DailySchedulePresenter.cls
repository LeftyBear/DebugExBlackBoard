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
    UseCase As App_LoadDailyScheduleUseCase
End Type

Private This As Member

Public Sub Inject(ByVal UseCase As App_LoadDailyScheduleUseCase)
    Set This.UseCase = UseCase
End Sub

