VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_CSVDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Dto")
Option Explicit

Private Type Member
    Fields() As String
End Type

Private This As Member

Public Sub Initialize(ByRef Values() As String)
    This.Fields = Values
End Sub

Public Property Get Field(ByVal Index As Long) As String
    Field = This.Fields(Index)
End Property

Public Property Get FieldCount() As Long
    FieldCount = UBound(This.Fields) - LBound(This.Fields) + 1
End Property

Public Function ToArray() As String()
    ToArray = This.Fields
End Function
