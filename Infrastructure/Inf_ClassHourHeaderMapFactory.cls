VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourHeaderMapFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal HeaderRows As VBA.Collection) As Inf_ClassHourHeaderMap
    Dim Result As Inf_ClassHourHeaderMap
    Set Result = New Inf_ClassHourHeaderMap
    Dim C As Long
    For C = 1 To HeaderRows.Count
        Dim Column As Inf_ClassHourColumn
        Set Column = CreateColumn(HeaderRows.Item(C))
        Result.Add C, Column
    Next
    Set Create = Result
End Function

Private Function CreateColumn(ByVal Header As String) As Inf_ClassHourColumn
    Dim Result As Inf_ClassHourColumn
    Set Result = New Inf_ClassHourColumn
    If 0 < VBA.InStr(1, Header, jpnHizuke) Then
        Result.RawDate = Header
    Else
        Result.RawID = Header
    End If
    Set CreateColumn = Result
End Function
