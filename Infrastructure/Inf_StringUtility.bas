Attribute VB_Name = "Inf_StringUtility"
'@Folder("Infrastructure.Service")
Option Explicit

Public Function JoinByBackSlash(ParamArray Strings() As Variant) As String
    Dim i As Long
    For i = LBound(Strings) To UBound(Strings)
        If VBA.Right$(Strings(i), 1) = charBackSlash Then
            Strings(i) = VBA.Left$(Strings(i), VBA.Len(Strings(i)) - 1)
        End If
    Next
    JoinByBackSlash = VBA.Join(Strings, charBackSlash)
End Function

Public Function GetHeaderRow(ByVal Rows As VBA.Collection) As String()
    Dim Result() As String
    ReDim Result(0 To Rows.Count - 1)
    Dim C As Long
    For C = 0 To Rows.Count - 1
        Result(C) = Rows.Item(C + 1)
    Next
    GetHeaderRow = Result
End Function
