VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    ClassHourFactory    As Dom_ClassHourFactory
    TimeTableFactory    As Dom_TimeTableFactory
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourFactory As Dom_ClassHourFactory, ByVal TimeTableFactory As Dom_TimeTableFactory)
    Set This.ClassHourFactory = ClassHourFactory
    Set This.TimeTableFactory = TimeTableFactory
End Sub

Public Function Create(ByVal Rows As VBA.Collection) As Dom_ClassHourYearAggregate
    Dim Yearly As Dom_ClassHourYearAggregate
    Set Yearly = New Dom_ClassHourYearAggregate
    Dim Grouped As Scripting.Dictionary
    Set Grouped = GroupByDate(Rows)
    Dim Key As Variant
    For Each Key In Grouped.Keys
        Dim Daily As Dom_ClassHourAggregate
        Set Daily = CreateDaily(Grouped.Item(Key))
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function

Private Function GroupByDate(ByVal Rows As VBA.Collection) As Scripting.Dictionary
    Dim Groups As Scripting.Dictionary
    Set Groups = New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_ClassHourRow
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

Private Function CreateDaily(ByVal Rows As VBA.Collection) As Dom_ClassHourAggregate
    Dim Aggregate As Dom_ClassHourAggregate
    Set Aggregate = New Dom_ClassHourAggregate
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_ClassHourRow
        Set Row = Rows.Item(i)
        If Row.Subject <> vbNullString Then
            Dim ClassHour As Dom_ClassHour
            Set ClassHour = This.ClassHourFactory.Create(Row)
            If Not ClassHour Is Nothing Then Aggregate.AddClassHour ClassHour
        ElseIf Row.Period <> 0 Then
            Dim TimeTable As Dom_TimeTable
            Set TimeTable = This.TimeTableFactory.Create(Row)
            If Not TimeTable Is Nothing Then Aggregate.AddTimeTable TimeTable
        End If
    Next
    Set CreateDaily = Aggregate
End Function
