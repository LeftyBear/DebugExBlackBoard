VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableMark"
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

Public Function Create(ByVal Value As String) As Dom_TimeTableMark
    Dim Result As Dom_TimeTableMark
    Set Result = New Dom_TimeTableMark
    Result.SetValueInternal Value
    Set Create = Result
End Function

Public Property Get Value() As String
    Value = This.Value
End Property

Friend Sub SetValueInternal(ByVal Value As String)
    This.Value = Value
End Sub
