VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentScalarModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit
Implements Dom_IEnrollmentColumn
Private Type Member
    Strategy As Dom_IEnrollmentStrategy
End Type

Private This As Member

Public Sub Initialize(ByVal Strategy As Dom_IEnrollmentStrategy)
    Set This.Strategy = Strategy
End Sub

Private Function Dom_IEnrollmentColumn_ApplyTo(ByVal Summary As Dom_EnrollmentSummary, ByVal Value As Long) As Dom_EnrollmentSummary
    Set Dom_IEnrollmentColumn_ApplyTo = This.Strategy.Apply(Summary, Value)
End Function
