VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolEventRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_ISchoolEventRepository
Private Type Member
    Persistence As Inf_SchoolEventPersistence
    EntList     As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_SchoolEventPersistence)
    Set This.Persistence = Persistence
    Set This.EntList = New Scripting.Dictionary
End Sub

Private Sub Dom_ISchoolEventRepository_Add(ByVal Entry As Dom_SchoolEvent)
    Dim Key As String
    Key = Entry.Value
    If This.EntList.Exists(Key) Then Err.Raise InfErrExistsKey, "Inf_SchoolEventRepository", "Exists Key: " & Key
    This.EntList.Add Key, Entry
End Sub

Private Function Dom_ISchoolEventRepository_Exsists(ByVal Name As String) As Boolean
    If This.EntList.Exists(Name) Then Dom_ISchoolEventRepository_Exsists = True
End Function

Private Function Dom_ISchoolEventRepository_GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Entry As Dom_SchoolEvent
    For Each Entry In This.EntList.Items
        Result.Add Entry
    Next
    Set Dom_ISchoolEventRepository_GetAll = Result
End Function

Private Function Dom_ISchoolEventRepository_ReadAll() As VBA.Collection
    Set Dom_ISchoolEventRepository_ReadAll = This.Persistence.ReadAll
End Function

Private Function Dom_ISchoolEventRepository_Replace(ByVal Name As String, ByVal NewEntry As Dom_SchoolEvent) As VBA.Collection
    If Not This.EntList.Exists(Name) Then Exit Function
    Set This.EntList.Item(Name) = NewEntry
    Set Dom_ISchoolEventRepository_Replace = Dom_ISchoolEventRepository_GetAll
End Function
