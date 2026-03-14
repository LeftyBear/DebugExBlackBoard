VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SpecialStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Names() As String
End Type

Private This As Member

Public Sub Initialize(ByRef Names() As String)
    This.Names = Names
End Sub

Public Property Get Count() As Long
    Count = UBound(This.Names)
End Property

Public Property Get Name(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise DomErrInvalidRange, "Dom_SpecialStream", "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.Count
    Name = This.Names(Index)
End Property

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= Me.Count)
End Function
