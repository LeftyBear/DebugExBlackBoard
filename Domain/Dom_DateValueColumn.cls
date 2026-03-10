VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DateValueColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit
Implements Dom_IValueObject
Private Type Member
    HeaderName  As String
    Value       As Date
End Type

Private This As Member

Public Sub Initialize(ByVal HeaderName As String, ByVal Value As Date)
    If HeaderName Like vbNullString Then Err.Raise DomErrNullString, , "HeaderName is null"
    This.HeaderName = HeaderName
    This.Value = Value
End Sub

Public Property Get Value() As Date
    Value = This.Value
End Property

Public Property Get HeaderName() As String
    HeaderName = This.HeaderName
End Property

Private Function Dom_IValueObject_IsEqual(ByVal Other As Dom_IValueObject) As Boolean
    If Not TypeOf Other Is Dom_DateValueColumn Then Exit Function
    Dim Target As Dom_DateValueColumn
    Set Target = Other
    Dom_IValueObject_IsEqual = (This.HeaderName = Target.HeaderName)
End Function
