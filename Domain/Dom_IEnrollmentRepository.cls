VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_IEnrollmentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.RepositoryInterface")
Option Explicit

Public Sub Add(ByVal Item As Dom_EnrollmentAggregate)
End Sub

Public Function FindByID(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
End Function

Public Function GetAll() As VBA.Collection
End Function
