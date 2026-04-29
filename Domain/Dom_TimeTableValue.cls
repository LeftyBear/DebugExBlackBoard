VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableValue"
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

Public Sub Initialize(ByVal Value As Long)
    This.Value = Value
End Sub

Public Property Get Value() As String
    Value = This.Value
End Property

Public Function Add(ByVal Other As Dom_TimeTableValue) As Dom_TimeTableValue
    Dim Result As Dom_TimeTableValue
    Set Result = New Dom_TimeTableValue
    Result.Initialize This.Value + Other.Value
    Set Add = Result
End Function

Public Function Equals(ByVal Other As Dom_TimeTableValue) As Boolean
    Equals = (This.Value = Other.Value)
End Function
