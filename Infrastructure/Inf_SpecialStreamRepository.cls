VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit
Implements Dom_ISpecialStreamRepository
Private Type Member
    Persistence As Inf_SpecialStreamPersistence
    NameIndex   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_SpecialStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.NameIndex Is Nothing Then Exit Sub
    Dim Rows As Inf_SpecialStreamRows
    Set Rows = This.Persistence.ReadAll
    BuildNameIndex Rows
End Sub

Private Sub BuildNameIndex(ByVal Rows As Inf_SpecialStreamRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SpecialStreamRow
        Set Row = Rows.Item(i)
        Dim Entry As Dom_SpecialStream
        Set Entry = New Dom_SpecialStream
        Entry.Initialize Row.Name
        This.NameIndex.Add Row.Name, Entry
    Next
End Sub

Private Sub ReplaceRow(ByVal Rows As Inf_SpecialStreamRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SpecialStreamRow
        Set Row = Rows.Item(i)
        If Not This.NameIndex.Exists(Row.Name) Then
            Rows.Remove i
        End If
    Next
    Dim Key As Variant
    For Each Key In This.NameIndex.Keys
        Dim Entry As Dom_SpecialStream
        Set Entry = This.NameIndex.Item(Key)
        For i = 1 To Rows.Count
            Set Row = Rows.Item(i)
            Row.Name = Entry.Value
        Next
    Next
End Sub

Private Function Dom_ISpecialStreamRepository_FindByName(ByVal Name As String) As Dom_SpecialStream
    EnsureLoaded
    Set Dom_ISpecialStreamRepository_FindByName = This.NameIndex.Item(Name)
End Function

Private Sub Dom_ISpecialStreamRepository_Replace(ByVal OldName As String, ByVal NewEntry As Dom_SpecialStream)
    EnsureLoaded
    If Not This.NameIndex.Exists(OldName) Then Exit Sub
    This.NameIndex.Remove OldName
    This.NameIndex.Add NewEntry.Value, NewEntry
End Sub

Private Sub Dom_ISpecialStreamRepository_Save()
    EnsureLoaded
    Dim Rows As Inf_SpecialStreamRows
    Set Rows = This.Persistence.ReadAll
    ReplaceRow Rows
    This.Persistence.WriteAll Rows
End Sub
