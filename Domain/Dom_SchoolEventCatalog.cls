VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolEventCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Catalog")
Option Explicit

Private Type Member
    Catalog As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Catalog = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal SchoolEvent As Dom_SchoolEvent)
    This.Catalog.Add SchoolEvent.Name, SchoolEvent
End Sub

Public Function Count() As Long
    Count = This.Catalog.Count
End Function

Public Function Exists(ByVal Key As String) As Boolean
    Exists = This.Catalog.Exists(Key)
End Function

Public Function Resolve(ByVal Key As String) As Dom_SchoolEvent
    If Not Me.Exists(Key) Then Err.Raise DomErrNotExistsItem, "Dom_SchoolEventCatalog", "存在しないアイテムを取得しようとしました。: " & vbCrLf & "未登録キー: " & Key
    Set Resolve = This.Catalog.Item(Key)
End Function

Public Function Names() As String()
    Dim Result() As String
    ReDim Result(This.Catalog.Count - 1)
    Dim Target As Variant
    For Each Target In This.Catalog.Keys
        Dim i As Long
        Result(i) = Target
        i = i + 1
    Next
    Names = Result
End Function

Public Function NameOf(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrNotExistsItem, "Dom_SubjectCatalog", "存在しないアイテムを取得しようとしました。: " & vbCrLf & "未登録インデックス: " & Index
    Dim Item As Variant
    For Each Item In This.Catalog.Items
        Dim SchoolEvent As Dom_SchoolEvent
        Set SchoolEvent = Item
        If SchoolEvent.Index = Index Then NameOf = SchoolEvent.Name: Exit For
    Next
End Function

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= This.Catalog.Count)
End Function
