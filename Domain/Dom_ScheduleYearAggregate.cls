VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleYearAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    Items As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Items = New VBA.Collection
End Sub

Public Sub Add(ByVal Daily As Dom_ScheduleAggregate)
    This.Items.Add Daily
End Sub

Public Function Count() As Long
    Count = This.Items.Count
End Function

Public Function Item(ByVal Index As Long) As Dom_ScheduleAggregate
    If Index < 1 Or This.Items.Count < Index Then Err.Raise DomErrNothingObject, "Dom_ScheduleYearAggregate", "ScheduleAggregateオブジェクトが必要です。"
    Set Item = This.Items.Item(Index)
End Function
