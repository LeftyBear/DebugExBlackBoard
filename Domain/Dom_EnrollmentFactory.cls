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

Public Function Create(ByVal ID As Dom_EnrollmentID, ByVal RawText As String) As Dom_Enrollment
    Dim Value As Dom_EnrollmentValue
    Set Value = Dom_ValueObjectFactory.CreateEnrollmentValue(RawText)
    Dim Result As Dom_Enrollment
    Set Result = New Dom_Enrollment
    Result.Initialize ID, Value
    Set Create = Result
End Function
