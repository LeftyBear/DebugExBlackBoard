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
        Dim Meta As Inf_ClassHourColumn
        Set Meta = CreateMeta(HeaderRows.Item(C))
        Result.Add C, Meta
    Next
    Set Create = Result
End Function

Private Function CreateMeta(ByVal Header As String) As Inf_ClassHourColumn
    Dim Meta As Inf_ClassHourColumn
    Set Meta = New Inf_ClassHourColumn
    If 0 < VBA.InStr(1, Header, jpnHizuke) Then
        Meta.RawDate = Header
    ElseIf 0 < VBA.InStr(1, Header, jpnIdou) Then
        Meta.RawTransfer = Header
    ElseIf 0 < VBA.InStr(1, Header, jpnBikou) Then
        Meta.RawRemarks = Header
    Else
        Meta.RawID = Header
    End If
    Set CreateMeta = Meta
End Function
