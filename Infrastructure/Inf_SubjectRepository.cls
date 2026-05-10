VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_ISubjectRepository
Private Type Member
    Persistence As Inf_SubjectPersistence
    EntList     As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_SubjectPersistence)
    Set This.Persistence = Persistence
    Set This.EntList = New Scripting.Dictionary
End Sub

Private Sub Dom_ISubjectRepository_Add(ByVal Entry As Dom_Subject)
    Dim Key As String
    Key = Entry.Name
    If This.EntList.Exists(Key) Then Err.Raise InfErrExistsKey, "Inf_SubjectRepository", "Exists Key: " & Key
    This.EntList.Add Key, Entry
End Sub

Private Function Dom_ISubjectRepository_Count() As Long
    Dom_ISubjectRepository_Count = This.EntList.Count
End Function

Private Function Dom_ISubjectRepository_Exists(ByVal Name As String) As Boolean
    Dom_ISubjectRepository_Exists = This.EntList.Exists(Name)
End Function

Private Function Dom_ISubjectRepository_FindByIndex(ByVal Index As Long) As Dom_Subject
    Dim Items() As Variant
    Items = This.EntList.Items
    Dim i As Long
    For i = 0 To This.EntList.Count - 1
        If i = Index - 1 Then
            Set Dom_ISubjectRepository_FindByIndex = Items(i)
        End If
    Next
End Function

Private Function Dom_ISubjectRepository_FindByName(ByVal Name As String) As Dom_Subject
    If This.EntList.Exists(Name) Then
        Set Dom_ISubjectRepository_FindByName = This.EntList.Item(Name)
    End If
End Function

Private Function Dom_ISubjectRepository_GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Entry As Dom_Subject
    For Each Entry In This.EntList.Items
        Result.Add Entry
    Next
    Set Dom_ISubjectRepository_GetAll = Result
End Function

Private Function Dom_ISubjectRepository_ReadAll() As VBA.Collection
    Set Dom_ISubjectRepository_ReadAll = This.Persistence.ReadAll
End Function

Private Function Dom_ISubjectRepository_Replace(ByVal Name As String, ByVal NewEntry As Dom_Subject) As VBA.Collection
    If Not This.EntList.Exists(Name) Then Exit Function
    Set This.EntList.Item(Name) = NewEntry
    Set Dom_ISubjectRepository_Replace = Dom_ISubjectRepository_GetAll
End Function
