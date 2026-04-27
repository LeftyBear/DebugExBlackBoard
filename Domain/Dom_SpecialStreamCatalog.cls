VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SpecialStreamCatalog"
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

Public Sub Add(ByVal SpecialStream As Dom_SpecialStream)
    This.Catalog.Add SpecialStream.Name, SpecialStream
End Sub

Public Function Count() As Long
    Count = This.Catalog.Count
End Function

Public Function Exists(ByVal Key As String) As Boolean
    Exists = This.Catalog.Exists(Key)
End Function

Public Function Item(ByVal Key As String) As Dom_SpecialStream
    If Not This.Catalog.Exists(Key) Then Err.Raise DomErrNotExistsItem, "Dom_SpecialStreamCatalog", "存在しないアイテムを取得しようとしました。: " & vbCrLf & "未登録キー: " & Key
    Set Item = This.Catalog.Item(Key)
End Function

Public Function Names() As String()
    Dim Result() As String
    ReDim Result(This.Catalog.Count - 1)
    Dim Item As Variant
    For Each Item In This.Catalog.Items
        Dim SpecialStream As Dom_SpecialStream
        Set SpecialStream = Item
        Dim i As Long
        Result(i) = SpecialStream.Name
        i = i + 1
    Next
    Names = Result
End Function

Public Function NameOf(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrNotExistsItem, "Dom_SpecialStreamCatalog", "存在しないアイテムを取得しようとしました。: " & vbCrLf & "未登録インデックス: " & Index
    Dim Item As Variant
    For Each Item In This.Catalog.Items
        Dim SpecialStream As Dom_SpecialStream
        Set SpecialStream = Item
        If SpecialStream.Index = Index Then NameOf = SpecialStream.Name: Exit For
    Next
End Function

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= This.Catalog.Count)
End Function
