VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal EntityGroups As VBA.Collection) As Dom_EnrollmentYearAggregate
    Dim Yearly As Dom_EnrollmentYearAggregate
    Set Yearly = New Dom_EnrollmentYearAggregate
    Dim EntityGroup As VBA.Collection
    For Each EntityGroup In EntityGroups
        Dim Daily As Dom_EnrollmentAggregate
        Set Daily = New Dom_EnrollmentAggregate
        Dim Enrollment As Dom_Enrollment
        For Each Enrollment In EntityGroup
            Daily.AddInitialEnrollment Enrollment
        Next
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function
