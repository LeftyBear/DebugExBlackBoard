VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_HeaderMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Public Function Map(ByVal Source As App_CSVDTO) As VBA.Collection
    Dim ColumnDTOGroups As VBA.Collection
    Set ColumnDTOGroups = New VBA.Collection
    Dim R As Long
    For R = 2 To Source.RowCount
        Dim ColumnDTOGroup As VBA.Collection
        Set ColumnDTOGroup = New VBA.Collection
        Dim C As Long
        For C = 1 To Source.ColumnCount
            Dim ColumnDTO As App_ColumnDTO
            Set ColumnDTO = New App_ColumnDTO
            ColumnDTO.Initialize Source.HeaderName(C), Source.Value(R, C)
            ColumnDTOGroup.Add ColumnDTO
        Next
        ColumnDTOGroups.Add ColumnDTOGroup
    Next
    Set Map = ColumnDTOGroups
End Function
