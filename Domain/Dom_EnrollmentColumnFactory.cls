VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentColumnFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal DTO As App_EnrollmentColumnDTO) As Dom_IEnrollmentColumn
    If DTO.IsScalar Then
        Set Create = CreateSpecial(DTO.ScalarCode)
        Exit Function
    End If
    Set Create = CreateStructured(DTO)
End Function

Private Function CreateStructured(ByVal DTO As App_EnrollmentColumnDTO) As Dom_IEnrollmentColumn
    Dim GradeColumn As Dom_GradeColumn
    Set GradeColumn = New Dom_GradeColumn
    GradeColumn.Initialize DTO.GradeNumber
    Dim ClassColumn As Dom_ClassColumn
    Set ClassColumn = New Dom_ClassColumn
    ClassColumn.Initialize DTO.ClassNumber
    Dim GenderColumn As Dom_GenderColumn
    Set GenderColumn = New Dom_GenderColumn
    GenderColumn.Initialize DTO.GenderCode
    Dim Strategy As Dom_EnrollmentCompositeStrategy
    Set Strategy = New Dom_EnrollmentCompositeStrategy
    Strategy.Initialize GradeColumn, ClassColumn, GenderColumn
    Dim Composite As Dom_EnrollmentCompositeColumn
    Set Composite = New Dom_EnrollmentCompositeColumn
    Composite.Initialize Strategy
    Set CreateStructured = Composite
End Function

Private Function CreateSpecial(ByVal Code As Long) As Dom_IEnrollmentColumn
    Dim ScalarColumn As Dom_ScalarColumn
    Set ScalarColumn = New Dom_ScalarColumn
    ScalarColumn.Initialize Code
    Dim Strategy As Dom_EnrollmentScalarStrategy
    Set Strategy = New Dom_EnrollmentScalarStrategy
    Strategy.Initialize ScalarColumn
    Dim Scalar As Dom_EnrollmentScalarModel
    Set Scalar = New Dom_EnrollmentScalarModel
    Scalar.Initialize Strategy
    Set CreateSpecial = Scalar
End Function
