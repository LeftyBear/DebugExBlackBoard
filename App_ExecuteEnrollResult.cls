VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteEnrollResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.DTO"
Option Explicit

Private Type Member
    Members As Dom_EnrollMemberList
    Summary As Dom_EnrollSummary
End Type

Private This As Member

Public Sub Initialize(ByVal Members As Dom_EnrollMemberList, ByVal Summary As Dom_EnrollSummary)
    Set This.Members = Members
    Set This.Summary = Summary
End Sub

Public Property Get Members() As Dom_EnrollMemberList
    Set Members = This.Members
End Property

Public Property Get Summary() As Dom_EnrollSummary
    Set Summary = This.Summary
End Property

