VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentRowsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_EnrollmentRows
    Dim Result As Inf_EnrollmentRows
    Set Result = New Inf_EnrollmentRows
    Dim MetaMapFactory As Inf_EnrollmentMetaMapFactory
    Set MetaMapFactory = New Inf_EnrollmentMetaMapFactory
    Dim MetaMap As Inf_EnrollmentMetaMap
    Set MetaMap = MetaMapFactory.Create(RawRows.GetHeaderRows)
    Dim RowFactory As Inf_EnrollmentRowFactory
    Set RowFactory = New Inf_EnrollmentRowFactory
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            Dim Meta As Inf_EnrollmentColumnMeta
            Set Meta = MetaMap.Item(C)
            Dim Row As Inf_EnrollmentRow
            Set Row = RowFactory.Create(Meta, RawRows.GetField(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
