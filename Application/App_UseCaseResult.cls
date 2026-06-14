VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_UseCaseResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit

Public Enum App_UseCaseResultTypeCode
    SuccessCode
    BusinessErrorCode
    SystemErrorCode
End Enum

Private Type Member
    TypeCode    As App_UseCaseResultTypeCode
    Message     As String
End Type

Private This As Member

Public Sub Initialize(ByVal TypeCode As App_UseCaseResultTypeCode, ByVal Message As String)
    This.TypeCode = TypeCode
    This.Message = Message
End Sub

Public Property Get TypeCode() As App_UseCaseResultTypeCode
    TypeCode = This.TypeCode
End Property

Public Property Get Message() As String
    Message = This.Message
End Property
