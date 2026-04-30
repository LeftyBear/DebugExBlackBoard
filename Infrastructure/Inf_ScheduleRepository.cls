VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_IScheduleRepository
Private Type Member
    ItemList As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ItemList = New Scripting.Dictionary
End Sub

Private Sub Dom_IScheduleRepository_Add(ByVal Item As Dom_ScheduleAggregate)
    This.ItemList.Add Item.DateID.ToKey, Item
End Sub

Private Function Dom_IScheduleRepository_FindByID(ByVal ID As Dom_ScheduleID) As Dom_Schedule
    If This.ItemList.Exists(ID.ToKey) Then
        Set Dom_IScheduleRepository_FindByID = This.ItemList.Item(ID.ToKey)
    End If
End Function

Private Function Dom_IScheduleRepository_GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Target As Variant
    For Each Target In This.ItemList.Items
        Result.Add Target
    Next
    Set Dom_IScheduleRepository_GetAll = Result
End Function
