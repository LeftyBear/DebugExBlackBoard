VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentScalarStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit
Implements Dom_IEnrollmentStrategy
Private Type Member
    Scalar As Dom_ScalarColumn
End Type

Private This As Member

Public Sub Initialize(ByVal Scalar As Dom_ScalarColumn)
    Set This.Scalar = Scalar
End Sub

Private Function Dom_IEnrollmentStrategy_Apply(ByVal Summary As Dom_EnrollmentSummary, ByVal Value As Long) As Dom_EnrollmentSummary

End Function
