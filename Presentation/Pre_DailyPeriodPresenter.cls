VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_DailyPeriodPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyPeriodPresenter

Public Function App_IDailyPeriodPresenter_CreateViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal MainStreamReadModel As App_MainStreamReadModel) As Variant()
    Dim Formatter As App_PeriodFormatter
    Set Formatter = This.Formatter
    Dim Table() As Variant
    Table = Formatter.Format(ExeKeyCode, ClassHourReadModels, MainStreamReadModel)
    App_IDailyPeriodPresenter_CreateViewModel = Table
End Function
