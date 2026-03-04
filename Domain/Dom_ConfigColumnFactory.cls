VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ConfigColumnFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit
Implements Dom_IEntityColumnFactory

Private Function Dom_IEntityColumnFactory_Create(ByVal ColumnDTO As App_ColumnDTO) As Dom_IEntityColumn
    Dim Scalar As Dom_ScalarColumn
    Set Scalar = New Dom_ScalarColumn
    Scalar.Initialize ColumnDTO.HeaderName, ColumnDTO.RawValue
    Set Dom_IEntityColumnFactory_Create = Scalar
End Function
