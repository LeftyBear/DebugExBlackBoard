VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_YearlyScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Private Type Member
    Factory As Dom_DailyScheduleFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_DailyScheduleFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByVal Rows As VBA.Collection) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Grouped As Scripting.Dictionary
    Set Grouped = GroupByDate(Rows)
    Dim Key As Variant
    For Each Key In Grouped.Keys
        Dim Daily As Dom_DailySchedule
        Set Daily = CreateDaily(Grouped.Item(Key))
        Result.Add Daily
    Next
    Set Create = Result
End Function

Private Function GroupByDate(ByVal Rows As VBA.Collection) As Scripting.Dictionary
    Dim Groups As Scripting.Dictionary
    Set Groups = New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_ScheduleRow
        Set Row = Rows.Item(i)
        Dim Key As String
        Key = CStr(Row.NormDate)
        If Not Groups.Exists(Key) Then
            Groups.Add Key, New VBA.Collection
        End If
        Dim Items As VBA.Collection
        Set Items = Groups.Item(Key)
        Items.Add Row
    Next
    Set GroupByDate = Groups
End Function

Private Function CreateDaily(ByVal Rows As VBA.Collection) As Dom_DailySchedule
    Dim Result As Dom_DailySchedule
    Set Result = New Dom_DailySchedule
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_ScheduleRow
        Set Row = Rows.Item(i)
        Dim Entry As Dom_Schedule
        Set Entry = This.Factory.Create(Row)
        If Not Entry Is Nothing Then Result.Add Entry
    Next
    Set CreateDaily = Result
End Function
