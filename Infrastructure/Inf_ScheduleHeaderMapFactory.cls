VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleHeaderMapFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal HeaderRows As VBA.Collection) As Inf_ScheduleHeaderMap
    Dim Result As Inf_ScheduleHeaderMap
    Set Result = New Inf_ScheduleHeaderMap
    Dim C As Long
    For C = 1 To HeaderRows.Count
        Dim Meta As Inf_ScheduleColumn
        Set Meta = CreateMeta(HeaderRows.Item(C))
        Result.Add C, Meta
    Next
    Set Create = Result
End Function

Private Function CreateMeta(ByVal Header As String) As Inf_ScheduleColumn
    Dim Meta As Inf_ScheduleColumn
    Set Meta = New Inf_ScheduleColumn
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
