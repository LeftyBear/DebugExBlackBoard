VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentCompositeStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit
Implements Dom_IEnrollmentStrategy
Private Type Member
    GradeColumn     As Dom_GradeColumn
    ClassColumn     As Dom_ClassColumn
    GenderColumn    As Dom_GenderColumn
End Type

Private This As Member

Public Sub Initialize(ByVal GradeColumn As Dom_GradeColumn, ByVal ClassColumn As Dom_ClassColumn, ByVal GenderColumn As Dom_GenderColumn)
    Set This.GradeColumn = GradeColumn
    Set This.ClassColumn = ClassColumn
    Set This.GenderColumn = GenderColumn
End Sub

Private Function Dom_IEnrollmentStrategy_Apply(ByVal Summary As Dom_EnrollmentSummary, ByVal Value As Long) As Dom_EnrollmentSummary
    
End Function
