VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_CSVLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Loader")
Option Explicit

Private Type Member
    RFCParser As App_CSVRFCParser
End Type
Private This As Member

Private Sub Class_Initialize()
    Set This.RFCParser = New App_CSVRFCParser
End Sub

Public Function Parse(ByRef Lines() As String) As VBA.Collection
    If UBound(Lines) < 1 Then Err.Raise errNoLines, , "No data lines"
    Dim Headers() As String
    Headers = This.RFCParser.ParseRecord(Lines(0))
    Dim RawRows As VBA.Collection
    Set RawRows = New VBA.Collection
    Dim i As Long
    For i = 1 To UBound(Lines)
        RawRows.Add BuildRow(Lines(i), Headers)
    Next
    Set Parse = RawRows
End Function

Private Function BuildRow(ByVal Line As String, ByRef Headers() As String) As VBA.Collection
    Dim Values() As String
    Values = This.RFCParser.ParseRecord(Line)
    If UBound(Values) <> UBound(Headers) Then Err.Raise errMismatchColumnCount, , "Column count mismatch"
    Dim RowsOfColumnDtos As VBA.Collection
    Set RowsOfColumnDtos = New VBA.Collection
    Dim i As Long
    For i = 0 To UBound(Headers)
        Dim DTO As App_EnrollmentColumnDTO
        Set DTO = New App_EnrollmentColumnDTO
        DTO.Initialize Headers(i), Values(i)
        RowsOfColumnDtos.Add DTO
    Next i
    Set BuildRow = RowsOfColumnDtos
End Function
