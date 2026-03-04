VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourColumnFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit
Implements Dom_IEntityColumnFactory

Private Function Dom_IEntityColumnFactory_Create(ByVal ColumnDTO As App_ColumnDTO) As Dom_IEntityColumn
    If 0 < VBA.InStr(1, ColumnDTO.HeaderName, charUnderScore) Then
        Dim Composite As Dom_EnrollmentCompositeStrategy
        Set Composite = New Dom_EnrollmentCompositeStrategy
        Set Dom_IEntityColumnFactory_Create = Composite.Create(ColumnDTO)
    Else
        Dim Scalar As Dom_ScalarColumn
        Set Scalar = New Dom_ScalarColumn
        Scalar.Initialize ColumnDTO.HeaderName, ColumnDTO.RawValue
        Set Dom_IEntityColumnFactory_Create = Scalar
    End If
End Function
