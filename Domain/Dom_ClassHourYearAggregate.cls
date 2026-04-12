VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourYearAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    DailyMap As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.DailyMap = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Daily As Dom_ClassHourAggregate)
    Dim Key As String
    Key = Daily.GetDateKey
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_ClassHourYearAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Sub UpdatePlanValue(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_ClassHourID, ByVal NewValue As Dom_ClassHourValue)
    Me.GetByDate(TargetDate).UpdatePlanValue ID, NewValue
End Sub

Public Sub UpdatePlanMark(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_TimeTableID, ByVal NewMark As Dom_TimeTableMark)
    Me.GetByDate(TargetDate).UpdatePlanMark ID, NewMark
End Sub

Public Sub UpdateExeValue(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_ClassHourID, ByVal NewValue As Dom_ClassHourValue)
    Me.GetByDate(TargetDate).UpdateExeValue ID, NewValue
End Sub

Public Sub UpdateExeMark(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_TimeTableID, ByVal NewMark As Dom_TimeTableMark)
    Me.GetByDate(TargetDate).UpdateExeMark ID, NewMark
End Sub

Public Function Count() As Long
    Count = This.DailyMap.Count
End Function

Public Function GetByDate(ByVal TargetDate As Dom_DateID) As Dom_ClassHourAggregate
    If Not This.DailyMap.Exists(TargetDate.ToKey) Then Err.Raise DomErrNothingObject, "Dom_ClassHourYearAggregate", "ClassHourAggregateオブジェクトが必要です。"
    Set GetByDate = This.DailyMap.Item(TargetDate.ToKey)
End Function
