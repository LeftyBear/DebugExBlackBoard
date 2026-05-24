VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolEventRowsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_SchoolEventRows
    Dim Result As Inf_SchoolEventRows
    Set Result = New Inf_SchoolEventRows
    Dim RowFactory As Inf_SchoolEventRowFactory
    Set RowFactory = New Inf_SchoolEventRowFactory
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            Dim Row As Inf_SchoolEventRow
            Set Row = RowFactory.Create(RawRows.GetField(1, C), RawRows.GetField(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
