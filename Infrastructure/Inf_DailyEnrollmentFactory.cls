VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyEnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Private Type Member
    Factory As Inf_EnrollmentFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Factory = New Inf_EnrollmentFactory
End Sub

Public Function Create(ByVal DailyRows As Inf_EnrollmentRows) As Dom_DailyEnrollments
    Dim Result As Dom_DailyEnrollments
    Set Result = New Dom_DailyEnrollments
    Dim i As Long
    For i = 1 To DailyRows.Count
        Dim Entry As Dom_Enrollment
        Set Entry = This.Factory.Create(DailyRows.Item(i))
        If Not Entry Is Nothing Then Result.Add Entry
    Next
    Set Create = Result
End Function
