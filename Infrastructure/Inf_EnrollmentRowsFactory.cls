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
    Dim MapFactory As New Inf_EnrollmentHeaderMapFactory
    Dim Map As Inf_EnrollmentHeaderMap
    Set Map = MapFactory.Create(RawRows.GetHeaderRows)
    Dim RowFactory As Inf_EnrollmentRowFactory
    Set RowFactory = New Inf_EnrollmentRowFactory
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            Dim Row As Inf_EnrollmentRow
            Set Row = RowFactory.Create(Map.Item(C), RawRows.GetField(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
