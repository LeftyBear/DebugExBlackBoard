VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
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

Public Sub Add(ByVal Daily As Dom_EnrollmentColumn)
    Dim Key As String
    Key = Daily.GetDateKey
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_EnrollmentAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Sub UpdateMainStreamValue(ByVal TargetDate As Date, ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    GetByDate(TargetDate).UpdateMainStream ID, NewValue
    ValidateTotal TargetDate, ID
End Sub

Public Sub UpdateSpecialStreamValue(ByVal TargetDate As Date, ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    GetByDate(TargetDate).UpdateSpecialStream ID, NewValue
    ValidateTotal TargetDate, ID
End Sub

Public Function Count() As Long
    Count = This.DailyMap.Count
End Function

Public Function GetByDate(ByVal TargetDate As Date) As Dom_EnrollmentColumn
    Dim Key As String
    Key = CDate(TargetDate)
    If This.DailyMap.Exists(Key) Then
        Set GetByDate = This.DailyMap.Item(Key)
    End If
End Function

Private Sub ValidateTotal(ByVal TargetDate As Date, ByVal ID As Dom_EnrollmentID)
    With New Dom_EnrollmentFilter
        Dim Filter As Dom_EnrollmentFilter
        Set Filter = .Create(ID.StreamType, ID.Grade, ID.ClassNo, ID.Gender)
    End With
    Dim Total As Long
    Total = GetByDate(TargetDate).GetTotal(Filter)
    If Total < 0 Then Err.Raise DomErrNegativeNumber, "Dom_EnrollmentAggregate", "合計は正の数になるはずです。"
End Sub
