VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentColumnDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Dto")
Option Explicit

Private Type Member
    GradeNumber As Long
    ClassNumber As Long
    GenderCode  As Long
    ScalarCode  As Long
    IsScalar    As Boolean
End Type

Private This As Member

Public Sub Initialize(ByVal GradeNumber As Long, ByVal ClassNumber As Long, ByVal GenderCode As Long, ByVal ScalarCode As Long, ByVal IsScalar As Boolean)
    This.GradeNumber = GradeNumber
    This.ClassNumber = ClassNumber
    This.GenderCode = GenderCode
    This.ScalarCode = ScalarCode
    This.IsScalar = IsScalar
End Sub

Public Property Get GradeNumber() As Long
    GradeNumber = This.GradeNumber
End Property

Public Property Get ClassNumber() As Long
    ClassNumber = This.ClassNumber
End Property

Public Property Get GenderCode() As Long
    GenderCode = This.GenderCode
End Property

Public Property Get IsScalar() As Boolean
    IsScalar = This.IsScalar
End Property

Public Property Get ScalarCode() As Long
    ScalarCode = This.ScalarCode
End Property
