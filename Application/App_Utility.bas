Attribute VB_Name = "App_Utility"
'@Folder "Application.Service"
Option Explicit
Option Private Module

Public Function ParseSearchKeywords(ByVal InputText As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Normalized As String
    Normalized = InputText
    Normalized = VBA.Replace$(Normalized, charWideSpace, charHalfSpace)
    Normalized = VBA.Replace$(Normalized, VBA.vbTab, charHalfSpace)
    Normalized = VBA.Replace$(Normalized, VBA.vbCr, charHalfSpace)
    Normalized = VBA.Replace$(Normalized, VBA.vbLf, charHalfSpace)
    Normalized = VBA.Trim$(Normalized)
    If VBA.Len(Normalized) = 0 Then
        Set ParseSearchKeywords = Result
        Exit Function
    End If
    Do While 0 < VBA.InStr(1, Normalized, charWideSpace)
        Normalized = VBA.Replace$(Normalized, charWideSpace, charHalfSpace)
    Loop
    Dim Parts() As String
    Parts = VBA.Split(Normalized, charHalfSpace)
    Dim Index As Long
    For Index = LBound(Parts) To UBound(Parts)
        If 0 < VBA.Len(Parts(Index)) Then Result.Add Parts(Index)
    Next
    Set ParseSearchKeywords = Result
End Function
