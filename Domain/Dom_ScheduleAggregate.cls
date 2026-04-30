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
    DateID      As Dom_DateID
    Entities    As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Entities = New VBA.Collection
End Sub

Public Sub Add(ByVal Entity As Dom_Schedule)
    If This.DateID Is Nothing Then Set This.DateID = Entity.DateID
    This.Entities.Add Entity
End Sub

Public Sub UpdateValue(ByVal ID As Dom_ScheduleID, ByVal NewValue As String)
    Dim Target As Dom_Schedule
    Set Target = FindByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ScheduleAggregate", "Scheduleオブジェクトが必要です。"
    Dim Value As Dom_ScheduleValue
    Set Value = New Dom_ScheduleValue
    Value.Initialize NewValue
    Target.ChangeValueInternal Value
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Function Count() As Long
    Count = This.Entities.Count
End Function

Public Function Item(ByVal Index As Long) As Dom_Schedule
    Set Item = This.Entities.Item(Index)
End Function

Private Function FindByID(ByVal ID As Dom_ScheduleID) As Dom_Schedule
    Dim i As Long
    For i = 1 To This.Entities.Count
        Dim Entity As Dom_Schedule
        Set Entity = This.Entities.Item(i)
        If Entity.ID.Equals(ID) Then
            Set FindByID = Entity: Exit Function
        End If
    Next
End Function
