VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal ColumnGroup As VBA.Collection) As Dom_EnrollmentYearAggregate
    Dim Yearly As Dom_EnrollmentYearAggregate
    Set Yearly = New Dom_EnrollmentYearAggregate
    Dim Daily As Dom_EnrollmentAggregate
    For Each Daily In ColumnGroup
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function
