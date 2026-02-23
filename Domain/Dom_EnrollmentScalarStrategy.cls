VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentScalarStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Strategy")
Option Explicit

Public Function Create(ByVal ColumnDTO As App_EnrollmentColumnDTO) As Dom_IEnrollmentColumn
    Dim Column As Dom_EnrollmentScalarColumn
    Set Column = New Dom_EnrollmentScalarColumn
    Column.Initialize ColumnDTO.HeaderName, ColumnDTO.RawValue
    Set Create = Column
End Function
