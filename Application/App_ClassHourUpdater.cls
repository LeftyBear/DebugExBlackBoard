VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Repository As App_ClassHourRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As App_ClassHourRepository)
    Set This.Repository = Repository
End Sub

Public Sub UpdateClassHourValue(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Dim Target As Dom_ClassHourAggregate
    Set Target = This.Repository.FindByID(TargetDate)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "App_ClassHourUpdater", "Object is nothing."
    Target.UpdateClassHourValue ID, NewValue
End Sub

Public Sub UpdateTimeTableValue(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Dim Target As Dom_ClassHourAggregate
    Set Target = This.Repository.FindByID(TargetDate)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "App_ClassHourUpdater", "Object is nothing."
    Target.UpdateTimeTableValue ID, NewValue
End Sub
