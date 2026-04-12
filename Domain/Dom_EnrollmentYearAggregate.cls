VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentYearAggregate"
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

Public Sub Add(ByVal Daily As Dom_EnrollmentAggregate)
    Dim Key As String
    Key = Daily.GetDateKey
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_EnrollmentYearAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Sub UpdateTransfer(ByVal TargetDate As Dom_DateID, ByVal NewValue As Dom_Transfer)
    Me.GetByDate(TargetDate).UpdateTransfer NewValue
End Sub

Public Sub UpdateRemarks(ByVal TargetDate As Dom_DateID, ByVal NewValue As Dom_Remarks)
    Me.GetByDate(TargetDate).UpdateRemarks NewValue
End Sub

Public Sub UpdateMainStream(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_EnrollmentID, ByVal NewValue As Dom_EnrollmentValue)
    Me.GetByDate(TargetDate).UpdateMainStream ID, NewValue
End Sub

Public Sub UpdateSpecialStream(ByVal TargetDate As Dom_DateID, ByVal ID As Dom_EnrollmentID, ByVal NewValue As Dom_EnrollmentValue)
    Me.GetByDate(TargetDate).UpdateSpecialStream ID, NewValue
End Sub

Public Function Count() As Long
    Count = This.DailyMap.Count
End Function

Public Function GetByDate(ByVal TargetDate As Dom_DateID) As Dom_EnrollmentAggregate
    If Not This.DailyMap.Exists(TargetDate.ToKey) Then Err.Raise DomErrNothingObject, "Dom_EnrollmentYearAggregate", "EnrollmentAggregateオブジェクトが必要です。"
    Set GetByDate = This.DailyMap.Item(TargetDate.ToKey)
End Function
