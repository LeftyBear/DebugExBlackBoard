VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Totalization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    SubjectName         As Dom_SubjectName
    SubjectTargetGrade  As Dom_SubjectTargetGrade
End Type

Private This As Member

Public Sub Initialize(ByVal SubjectName As Dom_SubjectName, ByVal SubjectTargetGrade As Dom_SubjectTargetGrade)
    This.SubjectName = SubjectName
    This.SubjectTargetGrade = SubjectTargetGrade
End Sub

Public Property Get SubjectCount() As Long
    SubjectCount = This.SubjectName.Count
End Property

Public Property Get SubjectName(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    SubjectName = This.SubjectName.GetValue(Index)
End Property

Public Property Get SubjectTargetGrade(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    SubjectTargetGrade = This.SubjectTargetGrade.GetValue(Index)
End Property

Public Function ToSubjectIndex(ByVal TargetName As String) As Long
    Dim Index As Long
    Index = GetIndex(TargetName)
    If Index = 0 Then Err.Raise AppErrNotFoundItem, , "Not found item: " & TargetName
    ToSubjectIndex = Index
End Function

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= This.SubjectName.Count)
End Function

Private Function SubjectContains(ByVal TargetName As String) As Boolean
    Dim i As Long
    For i = 1 To This.SubjectName.Count
        If This.SubjectName.GetValue(i) Like TargetName Then SubjectContains = True
    Next
End Function

Private Function GetIndex(ByVal TargetName As String) As Long
    Dim i As Long
    For i = 1 To This.SubjectName.Count
        If This.SubjectName.GetValue(i) Like TargetName Then GetIndex = i
    Next
End Function
