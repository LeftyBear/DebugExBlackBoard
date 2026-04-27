VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EntityFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Public Function CreateClassHour(ByVal DateID As Dom_DateID, ByVal ID As Dom_ClassHourID, ByVal Value As Dom_ClassHourValue) As Dom_ClassHour
    Dim Entity As Dom_ClassHour
    Set Entity = New Dom_ClassHour
    Entity.Initialize DateID, ID, Value
    Set CreateClassHour = Entity
End Function

Public Function CreateTransfer(ByVal DateID As Dom_DateID, ByVal Value As Boolean) As Dom_Transfer
    Dim Entity As Dom_Transfer
    Set Entity = New Dom_Transfer
    Entity.Initialize DateID, Value
    Set CreateTransfer = Entity
End Function

Public Function CreateRemarks(ByVal DateID As Dom_DateID, ByVal Value As String) As Dom_Remarks
    Dim Entity As Dom_Remarks
    Set Entity = New Dom_Remarks
    Entity.Initialize DateID, Value
    Set CreateRemarks = Entity
End Function

Public Function CreateEnrollment(ByVal DateID As Dom_DateID, ByVal ID As Dom_EnrollmentID, ByVal Value As Dom_EnrollmentValue) As Dom_Enrollment
    Dim Entity As Dom_Enrollment
    Set Entity = New Dom_Enrollment
    Entity.Initialize DateID, ID, Value
    Set CreateEnrollment = Entity
End Function
