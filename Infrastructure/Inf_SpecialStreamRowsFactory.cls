VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamRowsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_SpecialStreamRows
    Dim Result As Inf_SpecialStreamRows
    Set Result = New Inf_SpecialStreamRows
    Dim RowFactory As Inf_SpecialStreamRowFactory
    Set RowFactory = New Inf_SpecialStreamRowFactory
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Row As Inf_SpecialStreamRow
            Set Row = RowFactory.Create(RawRows.GetHeader(C), RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
