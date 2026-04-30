VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ScheduleColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Structure As Dom_SchoolStructure
End Type

Private This As Member

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Resolve(ByRef HeaderRow() As String) As Scripting.Dictionary
    Dim Map As Scripting.Dictionary
    Set Map = New Scripting.Dictionary
    Dim C As Long
    For C = LBound(HeaderRow) To UBound(HeaderRow)
        Dim ColumnName As String
        ColumnName = HeaderRow(C)
        Dim Meta As App_ScheduleColumnMeta
        Meta = ParseColumnMeta(ColumnName)
        Map.Add C, Meta
    Next
    Set Resolve = Map
End Function

Private Function ParseColumnMeta(ByVal ColumnName As String) As App_ScheduleColumnMeta
    Dim Meta As App_ScheduleColumnMeta
    Set Meta = New App_ScheduleColumnMeta
    If 0 < VBA.InStr(1, ColumnName, jpnHizuke) Then
        Meta.ColDate = ColumnName
    End If
    If This.Structure.IsSchoolEvent(ColumnName) Then
        Meta.EventName = ColumnName
    End If
    ParseColumnMeta = Meta
End Function
