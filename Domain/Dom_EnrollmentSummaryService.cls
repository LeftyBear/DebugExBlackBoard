VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentSummaryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Public Function Execute(ByVal Aggregate As Dom_EntityAggregate) As Dom_EnrollmentSummary
    Dim Summary As Dom_EnrollmentSummary
    Set Summary = New Dom_EnrollmentSummary
    Summary.Initialize
    Dim Record As Dom_EnrollmentRecord
'    For Each Record In Aggregate.EntityRows
'        Summary.Add Record.StreamType, Record.GradeNo, Record.ClassNo, Record.GenderCode, Record.Count
'    Next
    Set Execute = Summary
End Function
