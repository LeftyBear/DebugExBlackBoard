VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_PresenterResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit

Private Type Member
    TypeCode    As Pre_ViewResultTypePolicy
    Message     As String
End Type

Private This As Member

Public Sub Initialize(ByVal TypeCode As Pre_ViewResultTypePolicy, ByVal Message As String)
    This.TypeCode = TypeCode
    This.Message = Message
End Sub

Public Property Get TypeCode() As Pre_ViewResultTypePolicy
    TypeCode = This.TypeCode
End Property

Public Property Get Message() As String
    Message = This.Message
End Property
