VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScalarColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit
Implements Dom_IEntityColumn
Private Type Member
    HeaderName  As String
    Value       As String
End Type

Private This As Member

Public Sub Initialize(ByVal HeaderName As String, ByVal Value As String)
    If HeaderName Like vbNullString Then Err.Raise DomErrNullString, , "HeaderName is null"
    This.HeaderName = HeaderName
    This.Value = Value
End Sub

Private Property Get Dom_IEntityColumn_HeaderName() As String
    Dom_IEntityColumn_HeaderName = This.HeaderName
End Property

Private Property Get Dom_IEntityColumn_Value() As String
    Dom_IEntityColumn_Value = This.Value
End Property
