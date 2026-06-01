VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit
Implements Dom_ISubjectRepository
Private Type Member
    Factory     As Dom_ValueObjectFactory
    Persistence As Inf_SubjectPersistence
    NameIndex   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Inject(ByVal Factory As Dom_ValueObjectFactory, ByVal Persistence As Inf_SubjectPersistence)
    Set This.Factory = Factory
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.NameIndex Is Nothing Then Exit Sub
    Set This.NameIndex = New Scripting.Dictionary
    Dim Rows As Inf_SubjectRows
    Set Rows = This.Persistence.ReadAll
    BuildNameIndex Rows
End Sub

Private Sub BuildNameIndex(ByVal Rows As Inf_SubjectRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SubjectRow
        Set Row = Rows.Item(i)
        Dim Entry As Dom_Subject
        Set Entry = New Dom_Subject
        Dim Name As Dom_SubjectName
        Set Name = This.Factory.CreateSubjectName(Row.Name)
        Dim TargetGrade As Dom_SubjectTargetGrade
        Set TargetGrade = This.Factory.CreateSubjectTargetGrade(Row.TargetGrade)
        Dim Mark As Dom_SubjectMark
        Set Mark = This.Factory.CreateSubjectMark(Row.Mark)
        Entry.Initialize Name, TargetGrade, Mark
        This.NameIndex.Add Row.Name, Entry
    Next
End Sub

Private Sub ReplaceRow(ByRef Rows As Inf_SubjectRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SubjectRow
        Set Row = Rows.Item(i)
        If Not This.NameIndex.Exists(Row.Name) Then
            Rows.Remove i
        End If
    Next
    Dim Key As Variant
    For Each Key In This.NameIndex.Keys
        Dim Entry As Dom_Subject
        Set Entry = This.NameIndex.Item(Key)
        For i = 1 To Rows.Count
            Set Row = Rows.Item(i)
            Row.Name = Entry.Name.Value
            Row.TargetGrade = Entry.TargetGrade.Value
            Row.Mark = Entry.Mark.Value
        Next
    Next
End Sub

Private Function Dom_ISubjectRepository_FindByName(ByVal Name As String) As Dom_Subject
    EnsureLoaded
    If Not This.NameIndex.Exists(Name) Then Exit Function
    Set Dom_ISubjectRepository_FindByName = This.NameIndex.Item(Name)
End Function

Private Sub Dom_ISubjectRepository_Replace(ByVal OldName As String, ByVal NewEntry As Dom_Subject)
    EnsureLoaded
    If Not This.NameIndex.Exists(OldName) Then Exit Sub
    This.NameIndex.Remove OldName
    This.NameIndex.Add NewEntry.Name, NewEntry
End Sub

Private Sub Dom_ISubjectRepository_Save()
    EnsureLoaded
    Dim Rows As Inf_SubjectRows
    Set Rows = This.Persistence.ReadAll
    ReplaceRow Rows
    This.Persistence.WriteAll Rows
End Sub
