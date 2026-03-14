VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Names()         As String
    TargetGrades()  As String
    Marks()         As String
End Type

Private This As Member

Public Sub Initialize(ByRef Names() As String, ByRef TargetGrades() As String, ByRef Marks() As String)
    This.Names = Names
    This.TargetGrades = TargetGrades
    This.Marks = Marks
End Sub

Public Property Get Names() As String()
    Names = This.Names
End Property

Public Property Get TargetGrades() As String()
    TargetGrades = This.TargetGrades
End Property

Public Property Get Marks() As String()
    Marks = This.Marks
End Property

Public Function Count() As Long
    Count = UBound(This.Names)
End Function

Public Function Index(ByVal Name As String) As Long
    Dim i As Long
    For i = 1 To Me.Count
        If This.Names(i) Like Name Then
            Index = i: Exit Function
        End If
    Next
End Function

Public Function Name(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrInvalidRange, "Dom_Subject", "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.Count
    Name = This.Names(Index)
End Function

Public Function TargetGrade(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrInvalidRange, "Dom_Subject", "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.Count
    TargetGrade = This.TargetGrades(Index)
End Function

Public Function Mark(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrInvalidRange, "Dom_Subject", "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.Count
    Mark = This.Marks(Index)
End Function

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= Me.Count)
End Function
