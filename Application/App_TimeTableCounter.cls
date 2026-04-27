VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Filter As App_TimeTableFilter
End Type

Private This As Member

Public Sub Initialize(ByVal Filter As App_TimeTableFilter)
    Set This.Filter = Filter
End Sub

Public Function Count(ByVal Aggregate As Dom_ClassHourAggregate, ByVal Condition As App_TimeTableFilterCondition) As Long
    Dim TimeTables As VBA.Collection
    Set TimeTables = This.Filter.FilterByGrade(Aggregate, Condition)
    Dim i As Long
    For i = 1 To TimeTables.Count
        Dim TimeTable As Dom_TimeTable
        Set TimeTable = TimeTables.Item(i)
        If TimeTable.Value.Value <> vbNullString Then
            Count = Count + 1
        End If
    Next
End Function
