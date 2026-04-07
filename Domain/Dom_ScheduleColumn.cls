VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    NormDate    As Dom_DateID
    ItemList    As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal NormDate As Dom_DateID, ByVal ItemList As Scripting.Dictionary)
    This.NormDate = NormDate
    Set This.ItemList = ItemList
 End Sub

Public Sub UpdateSchedule(ByVal ID As Dom_ScheduleID, ByVal NewSchedule As String)
    Dim Schedule As Dom_Schedule
    Set Schedule = GetScheduleByID(ID)
    If Schedule Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ScheduleColumn", "IDをキーに持つオブジェクトが必要です。"
    Schedule.ChangeValueInternal NewSchedule
End Sub

Public Function GetDateKey() As String
    GetDateKey = This.NormDate.ToKey
End Function

Public Function GetValue(ByVal ID As Dom_ScheduleID) As String
    Dim Schedule As Dom_Schedule
    Set Schedule = GetScheduleByID(ID)
    If Schedule Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ScheduleColumn", "IDをキーに持つオブジェクトが必要です。"
    GetValue = Schedule.GetValue
End Function

Public Function GetIDs() As Variant()
    GetIDs = This.ItemList.Keys
End Function

Public Function GetValues() As Variant()
    GetValues = This.ItemList.Items
End Function

Private Function GetScheduleByID(ByVal ID As Dom_ScheduleID) As Dom_Schedule
    Dim Key As String
    Key = ID.ToKey
    If This.ItemList.Exists(Key) Then
        Set GetScheduleByID = This.ItemList.Item(Key)
    End If
End Function
