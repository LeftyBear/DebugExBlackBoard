VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_ISpecialStreamRepository
Private Type Member
    Persistence As Inf_SpecialStreamPersistence
    EntList     As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_SpecialStreamPersistence)
    Set This.Persistence = Persistence
    Set This.EntList = New Scripting.Dictionary
End Sub

Private Sub Dom_ISpecialStreamRepository_Add(ByVal Entry As Dom_SpecialStream)
    Dim Key As String
    Key = Entry.Value
    If This.EntList.Exists(Key) Then Err.Raise InfErrExistsKey, "Inf_SpecialStreamRepository", "Exists Key: " & Key
    This.EntList.Add Key, Entry
End Sub

Private Function Dom_ISpecialStreamRepository_Count() As Long
    Dom_ISpecialStreamRepository_Count = This.EntList.Count
End Function

Private Function Dom_ISpecialStreamRepository_FindByIndex(ByVal Index As Long) As Dom_SpecialStream
    If This.EntList.Exists(CStr(Index)) Then
        Set Dom_ISpecialStreamRepository_FindByIndex = This.EntList.Item(CStr(Index))
    End If
End Function

Private Function Dom_ISpecialStreamRepository_GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Entry As Dom_SpecialStream
    For Each Entry In This.EntList.Items
        Result.Add Entry
    Next
    Set Dom_ISpecialStreamRepository_GetAll = Result
End Function

Private Function Dom_ISpecialStreamRepository_ReadAll() As VBA.Collection
    Set Dom_ISpecialStreamRepository_ReadAll = This.Persistence.ReadAll
End Function

Private Function Dom_ISpecialStreamRepository_Replace(ByVal Name As String, ByVal NewEntry As Dom_SpecialStream) As VBA.Collection
    If Not This.EntList.Exists(Name) Then Exit Function
    Set This.EntList.Item(Name) = NewEntry
    Set Dom_ISpecialStreamRepository_Replace = Dom_ISpecialStreamRepository_GetAll
End Function
