Attribute VB_Name = "App_Util_Array"
'@Folder("Application.Utility")
Option Explicit

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
