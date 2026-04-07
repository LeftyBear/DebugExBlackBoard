VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleAggregate"
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

Public Sub Add(ByVal Daily As Dom_ScheduleColumn)
    Dim Key As String
    Key = Daily.GetDateKey
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_ScheduleAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Sub UpdateValue(ByVal TargetDate As Date, ByVal ID As Dom_ScheduleID, ByVal NewSchedule As String)
    Me.GetByDate(TargetDate).UpdateSchedule ID, NewSchedule
End Sub

Public Function Count() As Long
    Count = This.DailyMap.Count
End Function

Public Function GetByDate(ByVal TargetDate As Date) As Dom_ScheduleColumn
    Dim Key As String
    Key = CDate(TargetDate)
    If This.DailyMap.Exists(Key) Then
        Set GetByDate = This.DailyMap.Item(Key)
    End If
End Function
