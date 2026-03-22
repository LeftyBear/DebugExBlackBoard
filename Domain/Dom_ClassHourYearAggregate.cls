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
    Key = CStr(Daily.GetDate)
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_ClassHourYearAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Function GetByDate(ByVal TargetDate As Date) As Dom_ClassHourAggregate
    Dim Key As String
    Key = CDate(TargetDate)
    If This.DailyMap.Exists(Key) Then
        Set GetByDate = This.DailyMap.Item(Key)
    End If
End Function

Public Function GetDailyHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim Daily As Dom_ClassHourAggregate
    Daily = This.DailyMap.Item(TargetDate)
    GetDailyHourTotal = Daily.GetHourTotal(Filter)
End Function

Public Function GetWeeklyHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim TargetWeek As Long
    TargetWeek = Dom_Util_Date.GetSchoolWeek(TargetDate)
    Dim Dates() As Variant
    Dates = This.DailyMap.Keys
    Dim Years() As Variant
    Years = This.DailyMap.Years
    Dim i As Long
    For i = LBound(Years) To UBound(Years)
        If Dom_Util_Date.GetSchoolWeek(CDate(Dates(i))) = TargetWeek Then
            Dim Daily As Dom_ClassHourAggregate
            Set Daily = Years(i)
            Dim Total As Long
            Total = Total + Daily.GetHourTotal(Filter)
        End If
    Next
    GetWeeklyHourTotal = Total
End Function

Public Function GetMonthlyHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim TargetMonth As Long
    TargetMonth = VBA.Month(TargetDate)
    Dim Dates() As Variant
    Dates = This.DailyMap.Keys
    Dim Years() As Variant
    Years = This.DailyMap.Years
    Dim i As Long
    For i = LBound(Years) To UBound(Years)
        If VBA.Month(CDate(Dates(i))) = TargetMonth Then
            Dim Daily As Dom_ClassHourAggregate
            Set Daily = Years(i)
            Dim Total As Long
            Total = Total + Daily.GetHourTotal(Filter)
        End If
    Next
    GetMonthlyHourTotal = Total
End Function

Public Function GetCumulativeHourTotal(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim Dates() As Variant
    Dates = This.DailyMap.Keys
    Dim Years() As Variant
    Years = This.DailyMap.Years
    Dim i As Long
    For i = LBound(Years) To UBound(Years)
        If CDate(Dates(i)) <= TargetDate Then
            Dim Daily As Dom_ClassHourAggregate
            Set Daily = Years(i)
            Dim Total As Long
            Total = Total + Daily.GetHourTotal(Filter)
        Else
            Exit For
        End If
    Next
    GetCumulativeHourTotal = Total
End Function

Public Function GetMarkCount(ByVal TargetDate As Date, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim Daily As Dom_ClassHourAggregate
    Daily = This.DailyMap.Item(TargetDate)
    GetMarkCount = Daily.GetMarkCount(Filter)
End Function
