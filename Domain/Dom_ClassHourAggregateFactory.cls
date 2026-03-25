VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal EntityGroups As VBA.Collection) As Dom_ClassHourYearAggregate
    Dim Yearly As Dom_ClassHourYearAggregate
    Set Yearly = New Dom_ClassHourYearAggregate
    Dim EntityGroup As VBA.Collection
    For Each EntityGroup In EntityGroups
        Dim Daily As Dom_ClassHourAggregate
        Set Daily = New Dom_ClassHourAggregate
        Dim ClassHour As Dom_ClassHour
        For Each ClassHour In EntityGroup
            Daily.AddInitialClassHour ClassHour
        Next
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function
