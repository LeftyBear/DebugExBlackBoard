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

Public Function Create(ByVal ColumnDTO As App_EnrollmentColumnDTO) As Dom_IEnrollmentColumn
    If 0 < VBA.InStr(1, ColumnDTO.HeaderName, charUnderScore) Then
        Dim Composite As Dom_EnrollmentCompositeStrategy
        Set Composite = New Dom_EnrollmentCompositeStrategy
        Set Create = Composite.Create(ColumnDTO)
    Else
        Dim Scalar As Dom_EnrollmentScalarStrategy
        Set Scalar = New Dom_EnrollmentScalarStrategy
        Set Create = Scalar.Create(ColumnDTO)
    End If
End Function
