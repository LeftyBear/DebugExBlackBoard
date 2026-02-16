VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteClassHourResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.DTO"
Option Explicit

Private Type Member
    Members As Dom_ClassHourAggregate
    Summary As Dom_ClassHourSummary
End Type

Private This As Member

Public Sub Initialize(ByVal Members As Dom_ClassHourAggregate, ByVal Summary As Dom_ClassHourSummary)
    Set This.Members = Members
    Set This.Summary = Summary
End Sub

Public Property Get Members() As Dom_ClassHourAggregate
    Set Members = This.Members
End Property

Public Property Get Summary() As Dom_ClassHourSummary
    Set Summary = This.Summary
End Property

