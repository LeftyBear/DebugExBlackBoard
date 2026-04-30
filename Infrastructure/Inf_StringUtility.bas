Attribute VB_Name = "Inf_StringUtility"
Option Explicit

Public Function BuildToPath(ParamArray Strings() As Variant) As String
    BuildToPath = VBA.Join(Strings, charBackSlash)
End Function
