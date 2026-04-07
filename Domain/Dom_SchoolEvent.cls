VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Index   As Long
    Name    As String
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal Index As Long, ByVal Name As String)
    This.Index = Index
    This.Name = Name
End Sub

Public Function Create(ByVal Index As Long, ByVal Name As String) As Dom_SchoolEvent
    If Index < 1 Then Err.Raise DomErrInvalidValue, "Dom_SchoolEvent", "インデックスは正の数である必要があります。"
    If VBA.Len(Name) = 0 Then Err.Raise DomErrInvalidValue, "Dom_SchoolEvent", "イベント名は文字列である必要があります。"
    Dim Result As Dom_SchoolEvent
    Set Result = New Dom_SchoolEvent
    Result.InitializeInternal Index, Name
    Set Create = Result
End Function

Public Function Equals(ByVal Other As Dom_SchoolEvent) As Boolean
    If This.Index <> Other.Index Then Exit Function
    If This.Name <> Other.Name Then Exit Function
    Equals = True
End Function

Public Function ToKey() As String
    ToKey = This.Name
End Function

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Get Name() As String
    Name = This.Name
End Property
