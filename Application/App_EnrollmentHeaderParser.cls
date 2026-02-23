VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentHeaderParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Parser")
Option Explicit

Public Function Parse(ByVal CSVDTO As App_CSVDTO) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 0 To CSVDTO.FieldCount - 1
        Dim DTO As App_EnrollmentColumnDTO
        Set DTO = New App_EnrollmentColumnDTO
        DTO.Initialize i, CSVDTO.Field(i)
        Result.Add DTO
    Next
    Set Parse = Result
End Function
