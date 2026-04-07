VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Remarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As String
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal Value As String)
    This.Value = Value
End Sub

Public Function Create(ByVal Value As String) As Dom_Remarks
    Dim Result As Dom_Remarks
    Set Result = New Dom_Remarks
    Result.InitializeInternal Value
    Set Create = Result
End Function

Public Function Equals(ByVal Other As Dom_Remarks) As String
    If This.Value <> Other.Value Then Exit Function
    Equals = True
End Function

Public Property Get Value() As String
    Value = This.Value
End Property
