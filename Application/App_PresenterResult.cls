VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_PresenterResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Member
    ResultType  As App_ViewResultType
    Message     As String
End Type

Private This As Member

Public Sub Initialize(ByVal ResultType As App_ViewResultType, ByVal Message As String)
    This.ResultType = ResultType
    This.Message = Message
End Sub

Public Property Get ResultType() As App_ViewResultType
    ResultType = This.ResultType
End Property

Public Property Get Message() As String
    Message = This.Message
End Property
