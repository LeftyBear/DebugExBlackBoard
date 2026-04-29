VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SpecialStream"
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

Public Sub Initialize(ByVal Index As Long, ByVal Name As String)
    If Index < 1 Then Err.Raise DomErrNegativeNumber, "Dom_SpecialStream", "インデックスは1以上である必要があります。"
    If VBA.Len(Name) = 0 Then Err.Raise DomErrNullString, "Dom_SpecialStream", "集計名は文字列である必要があります。"
    This.Index = Index
    This.Name = Name
End Sub

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Function Equals(ByVal Other As Dom_SpecialStream) As Boolean
    If This.Index <> Other.Index Then Exit Function
    If This.Name <> Other.Name Then Exit Function
    Equals = True
End Function
