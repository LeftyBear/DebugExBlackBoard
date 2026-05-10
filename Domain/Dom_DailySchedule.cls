VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DailySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID  As Dom_DateID
    Entries As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Entries = New VBA.Collection
End Sub

Public Sub Add(ByVal Entry As Dom_Schedule)
    If This.DateID Is Nothing Then Set This.DateID = Entry.DateID
    This.Entries.Add Entry
End Sub

Public Sub UpdateValue(ByVal ID As Dom_ScheduleID, ByVal NewValue As String)
    Dim Target As Dom_Schedule
    Set Target = FindByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_DailySchedule", "Scheduleオブジェクトが必要です。"
    Dim Value As Dom_ScheduleValue
    Set Value = New Dom_ScheduleValue
    Value.Initialize NewValue
    Target.ChangeValueInternal Value
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Function Count() As Long
    Count = This.Entries.Count
End Function

Public Function Item(ByVal Index As Long) As Dom_Schedule
    Set Item = This.Entries.Item(Index)
End Function

Private Function FindByID(ByVal ID As Dom_ScheduleID) As Dom_Schedule
    Dim i As Long
    For i = 1 To This.Entries.Count
        Dim Entry As Dom_Schedule
        Set Entry = This.Entries.Item(i)
        If Entry.ID.Equals(ID) Then
            Set FindByID = Entry: Exit Function
        End If
    Next
End Function
