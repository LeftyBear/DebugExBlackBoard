VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_HeaderParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Parser")
Option Explicit

Public Function Parse(ByVal Source As App_CSVDTO) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim R As Long
    For R = 1 To Source.RowCount
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = New VBA.Collection
        Dim C As Long
        For C = 1 To Source.ColumnCount
            Dim DTO As App_ColumnDTO
            Set DTO = New App_ColumnDTO
            DTO.Initialize Source.HeaderName(C), Source.Value(R, C)
            ColumnDTOs.Add DTO
        Next
        Result.Add ColumnDTOs
    Next
    Set Parse = Result
End Function
