VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit

Private Type Member
    ItemList As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ItemList = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Item As Dom_ClassHourAggregate)
    This.ItemList.Add Item.DateID.ToKey, Item
End Sub

Public Function FindByID(ByVal ID As Dom_DateID) As Dom_ClassHourAggregate
    If This.ItemList.Exists(ID.ToKey) Then
        Set FindByID = This.ItemList.Item(ID.ToKey)
    End If
End Function

Public Function GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Target As Variant
    For Each Target In This.ItemList.Items
        Result.Add Target
    Next
    Set GetAll = Result
End Function
