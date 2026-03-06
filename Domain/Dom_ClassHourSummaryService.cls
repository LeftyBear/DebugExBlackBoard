VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourSummaryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Public Function Execute(ByVal Aggregate As Dom_EntityAggregate) As Dom_ClassHourSummary
    Dim Summary As Dom_ClassHourSummary
    Set Summary = New Dom_ClassHourSummary
    Summary.Initialize
    Dim Rec As Dom_ClassHourRecord
'    For Each Rec In Aggregate.EntityRows
'        Summary.Add Rec.GradeNo, Rec.ClassNo, Rec.GenderCode, Rec.Count
'    Next
    Set Execute = Summary
End Function
