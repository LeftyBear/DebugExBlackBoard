VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_IDailyClassHourPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Interface")
Option Explicit

Public Function CreatePlanViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Pre_TableViewModel
End Function

Public Function CreateExeViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels, ByVal Totals As App_ClassHourTotals, ByVal GradeCount As Long) As Pre_TableViewModel
End Function
