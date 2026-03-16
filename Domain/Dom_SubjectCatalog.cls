VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SubjectCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Catalog As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Catalog = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Subject As Dom_Subject)
    This.Catalog.Add Subject.Name, Subject
End Sub

Public Function Exists(ByVal Name As String) As Boolean
    Exists = This.Catalog.Exists(Name)
End Function

Public Function Count() As Long
    Count = This.Catalog.Count
End Function

Public Function Names() As String()
    Dim Result() As String
    ReDim Result(This.Catalog.Count - 1)
    Dim Item As Variant
    For Each Item In This.Catalog.Items
        Dim Subject As Dom_Subject
        Set Subject = Item
        Dim i As Long
        Result(i) = Subject.Name
        i = i + 1
    Next
    Names = Result
End Function

Public Function Resolve(ByVal Name As String) As Dom_Subject
    If Not This.Catalog.Exists(Name) Then Exit Function
    Set Resolve = This.Catalog.Item(Name)
End Function

Public Function NameOf(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrNotExistsItem, "Dom_SubjectCatalog", "存在しないアイテムを取得しようとしました。: " & vbCrLf & "未登録インデックス: " & Index
    Dim Item As Variant
    For Each Item In This.Catalog.Items
        Dim Subject As Dom_Subject
        Set Subject = Item
        If Subject.Index = Index Then NameOf = Subject.Name: Exit For
    Next
End Function

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= This.Catalog.Count)
End Function
