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

Public Function Build(ByVal Aggregate As Dom_EnrollmentAggregate) As Dom_EnrollmentSummary
    Dim Summary As Dom_EnrollmentSummary
    Set Summary = New Dom_EnrollmentSummary
    Summary.Initialize
    Dim Rec As Dom_EnrollmentRecord
    For Each Rec In Aggregate.Records
        Summary.Add Rec.StreamType, Rec.GradeNo, Rec.ClassNo, Rec.GenderCode, Rec.Count
    Next
    Set Build = Summary
End Function
