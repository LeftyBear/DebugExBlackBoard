VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_YearlyClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit
Private Type Member
    Factory As Inf_DailyClassHourFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Factory = New Inf_DailyClassHourFactory
End Sub

Public Function Create(ByVal Rows As Inf_ClassHourRows) As Dom_YearlyClassHour
    Dim Result As Dom_YearlyClassHour
    Set Result = New Dom_YearlyClassHour
    Dim DateIndex As Scripting.Dictionary
    Set DateIndex = GroupByDate(Rows)
    Dim Key As Variant
    For Each Key In DateIndex.Keys
        Dim Daily As Dom_DailyClassHour
        Set Daily = CreateDaily(CDate(Key), DateIndex.Item(Key))
        Result.Add Daily
    Next
    Set Create = Result
End Function

Private Function GroupByDate(ByVal Rows As Inf_ClassHourRows) As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_ClassHourRow
        Set Row = Rows.Item(i)
        Dim Key As String
        Key = CStr(Row.NormDate)
        If Not Result.Exists(Key) Then
            Result.Add Key, New VBA.Collection
        End If
        Dim Items As VBA.Collection
        Set Items = Result.Item(Key)
        Items.Add Row
    Next
    Set GroupByDate = Result
End Function

Private Function CreateDaily(ByVal TargetDate As Date, ByVal Rows As VBA.Collection) As Dom_DailyClassHour
    Dim Result As Dom_DailyClassHour
    Set Result = New Dom_DailyClassHour
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_ClassHourRow
        Set Row = Rows.Item(i)
        Dim Entry As Dom_ClassHour
        Set Entry = This.Factory.Create(TargetDate, Row)
        If Not Entry Is Nothing Then Result.Add Entry
    Next
    Set CreateDaily = Result
End Function
