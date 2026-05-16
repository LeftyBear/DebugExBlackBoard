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
