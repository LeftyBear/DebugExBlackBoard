VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal ID As Dom_EnrollmentID, ByVal RawValue As Long) As Dom_Enrollment
    '@Ignore SelfAssignedDeclaration
    Dim VOFactory As New Dom_ValueObjectFactory
    Dim Value As Dom_EnrollmentValue
    Set Value = VOFactory.CreateEnrollmentValue(RawValue)
    Dim Result As Dom_Enrollment
    Set Result = New Dom_Enrollment
    Result.Initialize ID, Value
    Set Create = Result
End Function
