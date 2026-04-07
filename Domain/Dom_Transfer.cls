VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As Boolean
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal Value As Boolean)
    This.Value = Value
End Sub

Public Function Create(ByVal Value As Boolean) As Dom_Transfer
    Dim Result As Dom_Transfer
    Set Result = New Dom_Transfer
    Result.InitializeInternal Value
    Set Create = Result
End Function

Public Function Equals(ByVal Other As Dom_Transfer) As Boolean
    If This.Value <> Other.Value Then Exit Function
    Equals = True
End Function

Public Property Get Value() As Boolean
    Value = This.Value
End Property
