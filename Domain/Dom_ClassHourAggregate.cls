VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregate"
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

Public Sub Add(ByVal Daily As Dom_ClassHourColumn)
    Dim Key As String
    Key = Daily.GetDateKey
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_ClassHourAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Sub UpdatePlanValue(ByVal TargetDate As Date, ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Me.GetByDate(TargetDate).UpdatePlanValue ID, NewValue
    ValidateTotal TargetDate, ID
End Sub

Public Sub UpdatePlanMark(ByVal TargetDate As Date, ByVal ID As Dom_ClassHourID, ByVal NewMark As String)
    Me.GetByDate(TargetDate).UpdatePlanMark ID, NewMark
    ValidateTotal TargetDate, ID
End Sub

Public Sub UpdateExeValue(ByVal TargetDate As Date, ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Me.GetByDate(TargetDate).UpdateExeValue ID, NewValue
    ValidateTotal TargetDate, ID
End Sub

Public Sub UpdateExeMark(ByVal TargetDate As Date, ByVal ID As Dom_ClassHourID, ByVal NewMark As String)
    Me.GetByDate(TargetDate).UpdateExeMark ID, NewMark
    ValidateTotal TargetDate, ID
End Sub

Public Function Count() As Long
    Count = This.DailyMap.Count
End Function

Public Function GetByDate(ByVal TargetDate As Date) As Dom_ClassHourColumn
    Dim Daily As Dom_ClassHourColumn
    Daily = Me.GetByDate(TargetDate)
    If Not Daily Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "オブジェクトが見つかりません。"
    Set GetByDate = Daily
End Function

Public Function GetWeeklyHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim TargetWeeks As Long
    TargetWeeks = Dom_Util_Date.GetSchoolWeek(TargetDate)
    Dim DateKey As Variant
    For Each DateKey In This.DailyMap.Keys
        If Dom_Util_Date.GetSchoolWeek(CDate(DateKey)) = TargetWeeks Then
            Dim Daily As Dom_ClassHourColumn
            Set Daily = This.DailyMap.Item(DateKey)
            Dim Total As Long
            Total = Total + Daily.GetHourTotal(Filter)
        End If
    Next
    GetWeeklyHourTotal = Total
End Function

Public Function GetMonthlyHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim TargetMonth As Long
    TargetMonth = VBA.Month(TargetDate)
    Dim DateKey As Variant
    For Each DateKey In This.DailyMap.Keys
        If VBA.Month(CDate(DateKey)) = TargetMonth Then
            Dim Daily As Dom_ClassHourColumn
            Set Daily = This.DailyMap.Item(DateKey)
            Dim Total As Long
            Total = Total + Daily.GetHourTotal(Filter)
        End If
    Next
    GetMonthlyHourTotal = Total
End Function

Public Function GetCumulativeHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim DateKey As Variant
    For Each DateKey In This.DailyMap.Keys
        If CDate(DateKey) <= TargetDate Then
            Dim Daily As Dom_ClassHourColumn
            Set Daily = This.DailyMap.Item(DateKey)
            Dim Total As Long
            Total = Total + Daily.GetHourTotal(Filter)
        End If
    Next
    GetCumulativeHourTotal = Total
End Function

Private Sub ValidateTotal(ByVal TargetDate As Date, ByVal ID As Dom_ClassHourID)
    Dim Daily As Dom_ClassHourColumn
    Set Daily = Me.GetByDate(TargetDate)
    With New Dom_ClassHourFilter
        Dim ClassHourFilter As Dom_ClassHourFilter
        Set ClassHourFilter = .Create(ID.ClassHourType, Nothing, ID.Grade)
    End With
    Dim HourTotal As Long
    HourTotal = Daily.GetHourTotal(ClassHourFilter)
    With New Dom_TimeTableFilter
        Dim TimeTableFilter As Dom_TimeTableFilter
        Set TimeTableFilter = .Create(ID.ClassHourType, Nothing, ID.Grade)
    End With
    Dim MarkCount As Long
    MarkCount = Daily.GetMarkCount(TimeTableFilter)
    If HourTotal = MarkCount Then Exit Sub
    Err.Raise DomErrUnmatch, "Dom_ClassHourAggregate", ID.ClassHourType.ToKey & charHalfSpace & ID.Grade.ToKey & "の時数と時間割を一致させてください。"
End Sub
