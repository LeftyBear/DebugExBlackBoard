VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Names           As VBA.Collection
    TargetGrades    As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Names As VBA.Collection, ByVal TargetGrades As VBA.Collection)
    Set This.Names = Names
    Set This.TargetGrades = TargetGrades
End Sub

Public Property Get Count() As Long
    Count = This.Names.Count
End Property

Public Property Get Name(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    Name = This.Names.Item(Index)
End Property

Public Property Get TargetGrade(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    TargetGrade = This.TargetGrades.Item(Index)
End Property

Public Function Names(ByVal Index As Long) As String()
    Dim Result() As String
    ReDim Result(This.Names.Count - 1)
    Dim i As Long
    For i = 0 To This.Names.Count - 1
        Result(i) = This.Names.Item(i + 1)
    Next
    Names = Result
End Function

Public Function TargetGrades(ByVal Index As Long) As String()
    Dim Result() As String
    ReDim Result(This.TargetGrades.Count - 1)
    Dim i As Long
    For i = 0 To This.TargetGrades.Count - 1
        Result(i) = This.TargetGrades.Item(i + 1)
    Next
    TargetGrades = Result
End Function

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= This.Names.Count)
End Function
