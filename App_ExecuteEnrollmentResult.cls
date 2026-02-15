VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteEnrollmentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.DTO"
Option Explicit

Private Type Member
    Members As Dom_EnrollmentList
    Summary As Dom_EnrollmentSummary
End Type

Private This As Member

Public Sub Initialize(ByVal Members As Dom_EnrollmentList, ByVal Summary As Dom_EnrollmentSummary)
    Set This.Members = Members
    Set This.Summary = Summary
End Sub

Public Property Get Members() As Dom_EnrollmentList
    Set Members = This.Members
End Property

Public Property Get Summary() As Dom_EnrollmentSummary
    Set Summary = This.Summary
End Property

