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
    Key = CStr(Daily.GetDate)
    If This.DailyMap.Exists(Key) Then Err.Raise DomErrNotUnique, "Dom_EnrollmentYearAggregate", "キーは一意である必要があります。" & vbCrLf & "重複している日付キー: " & Key
    This.DailyMap.Add Key, Daily
End Sub

Public Function GetByDate(ByVal TargetDate As Date) As Dom_EnrollmentAggregate
    Dim Key As String
    Key = CDate(TargetDate)
    If This.DailyMap.Exists(Key) Then
        Set GetByDate = This.DailyMap.Item(Key)
    End If
End Function

Public Function GetMaleTotal(ByVal TargetDate As Date, ByVal Filter As Dom_EnrollmentFilter) As Long
    Dim Daily As Dom_EnrollmentAggregate
    Set Daily = This.DailyMap.Item(TargetDate)
    GetMaleTotal = Daily.GetValueTotal(Filter)
End Function

Public Function GetFemaleTotal(ByVal TargetDate As Date, ByVal Filter As Dom_EnrollmentFilter) As Long
    Dim Daily As Dom_EnrollmentAggregate
    Set Daily = This.DailyMap.Item(TargetDate)
    GetFemaleTotal = Daily.GetValueTotal(Filter)
End Function
