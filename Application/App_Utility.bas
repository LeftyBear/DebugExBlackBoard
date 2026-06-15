Attribute VB_Name = "App_Utility"
'@Folder "Application.Service"
Option Explicit
Option Private Module
Private Const WIDE_SPACE As String = "Å@"

Public Function ParseSearchKeywords(ByVal InputText As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Const HALF_SPACE As String = " "
    Dim Normalized As String
    Normalized = InputText
    Normalized = VBA.Replace$(Normalized, WIDE_SPACE, HALF_SPACE)
    Normalized = VBA.Replace$(Normalized, VBA.vbTab, HALF_SPACE)
    Normalized = VBA.Replace$(Normalized, VBA.vbCr, HALF_SPACE)
    Normalized = VBA.Replace$(Normalized, VBA.vbLf, HALF_SPACE)
    Normalized = VBA.Trim$(Normalized)
    If VBA.Len(Normalized) = 0 Then
        Set ParseSearchKeywords = Result
        Exit Function
    End If
    Do While 0 < VBA.InStr(1, Normalized, WIDE_SPACE)
        Normalized = VBA.Replace$(Normalized, WIDE_SPACE, HALF_SPACE)
    Loop
    Dim Parts() As String
    Parts = VBA.Split(Normalized, HALF_SPACE)
    Dim Index As Long
    For Index = LBound(Parts) To UBound(Parts)
        If 0 < VBA.Len(Parts(Index)) Then Result.Add Parts(Index)
    Next
    Set ParseSearchKeywords = Result
End Function
