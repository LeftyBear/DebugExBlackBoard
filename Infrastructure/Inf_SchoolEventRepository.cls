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
    Factory     As Dom_ValueObjectFactory
    Persistence As Inf_SchoolEventPersistence
    NameIndex   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory, ByVal Persistence As Inf_SchoolEventPersistence)
    Set This.Factory = Factory
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.NameIndex Is Nothing Then Exit Sub
    Set This.NameIndex = New Scripting.Dictionary
    Dim RawRows As Inf_RawRows
    Set RawRows = This.Persistence.ReadAll
    BuildNameIndex RawRows
End Sub

Private Sub BuildNameIndex(ByVal RawRows As Inf_RawRows)
    Dim Factory As Inf_SchoolEventRowFactory
    Set Factory = New Inf_SchoolEventRowFactory
    Dim R As Long
    For R = 2 To RawRows.RecordsCount
        Dim C As Long
        For C = 1 To RawRows.FieldsCount
            Dim Row as Inf_SchoolEventRow
            Set Row = Factory.Create(RawRows.GetField(1, C),RawRows.GetField(R, C))
            Dim Entry As Dom_SchoolEvent
            Set Entry = New Dom_SchoolEvent
            Entry.Initialize Row.Name
            This.NameIndex.Add Entry.Value, Entry
        Next
    Next
End Sub

Private Sub ReplaceRow(ByRef Rows As Inf_SchoolEventRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SchoolEventRow
        Set Row = Rows.Item(i)
        If Not This.NameIndex.Exists(Row.Name) Then
            Rows.Remove i
        End If
    Next
    Dim Key As Variant
    For Each Key In This.NameIndex.Keys
        Dim Entry As Dom_SchoolEvent
        Set Entry = This.NameIndex.Item(Key)
        For i = 1 To Rows.Count
            Set Row = Rows.Item(i)
            Row.Name = Entry.Value
        Next
    Next
End Sub

Private Function Dom_ISchoolEventRepository_FindByName(ByVal Name As String) As Dom_SchoolEvent
    EnsureLoaded
    If This.NameIndex.Exists(Name) Then
        Set Dom_ISchoolEventRepository_FindByName = This.NameIndex.Item(Name)
    End If
End Function

Private Sub Dom_ISchoolEventRepository_Replace(ByVal OldName As String, ByVal NewEntry As Dom_SchoolEvent)
    EnsureLoaded
    If Not This.NameIndex.Exists(OldName) Then Exit Sub
    This.NameIndex.Remove OldName
    This.NameIndex.Add NewEntry.Value, NewEntry
End Sub

Private Sub Dom_ISchoolEventRepository_Save()
    EnsureLoaded
    Dim Rows As Inf_SchoolEventRows
    Set Rows = This.Persistence.ReadAll
    ReplaceRow Rows
    This.Persistence.WriteAll Rows
End Sub
