Attribute VB_Name = "App_StringArrayUtility"
'@Folder("Application.Service")
Option Explicit

Public Function GetHeaderRow(ByRef Source() As String) As String()
    Dim Result() As String
    ReDim Result(UBound(Source, 2) - 1)
    Dim C As Long
    For C = LBound(Source, 2) To UBound(Source, 2)
        Result(C - 1) = Source(1, C)
    Next
    GetHeaderRow = Result
End Function

Public Function IsEmpty(ByRef Source() As String) As Boolean
    If Not VBA.IsArray(Source) Then IsEmpty = True: Exit Function
    On Error Resume Next
    Dim Lower As Long
    Lower = LBound(Source)
    Dim Upper As Long
    Upper = UBound(Source)
    If Err.Number <> 0 Then
        IsEmpty = True
    Else
        IsEmpty = (Upper < Lower)
    End If
    On Error GoTo 0
End Function
