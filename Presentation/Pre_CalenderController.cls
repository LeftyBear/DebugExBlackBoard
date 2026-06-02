VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_CalenderController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Controller")
Option Explicit

Private Type Member
    Calender    As Pre_CalenderPresenter
    Schedule    As Pre_DailySchedulePresenter
    CurrentYear As Long
End Type

Private This As Member

Public Sub Initialize(ByVal Calender As Pre_CalenderPresenter, ByVal Schedule As Pre_DailySchedulePresenter)
    Set This.Calender = Calender
    Set This.Schedule = Schedule
End Sub

Public Sub LoadOf(ByVal SchoolYear As Long)
    
End Sub
